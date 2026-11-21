#!/usr/bin/env python3
"""
TprONE_ini_to_sql.py - Convert a TPR configuration INI file into a Tibero SQL script.
Generates a tbsql collection script from a TPR (Tibero Performance Repository) INI file.

Made By, yeonhong.min@gmail.com
Version 1.0 2025.11.21
"""

import io
import sys
import os
import re
import argparse
from collections import OrderedDict

# Debug mode flag (set to True to print debug output)
debug = False

def detect_encoding(path):
    """
    Detect the character encoding of an INI file.

    Returns 'utf-8' if the entire file decodes cleanly as UTF-8, otherwise 'cp949'
    (default Korean Windows encoding / superset of EUC-KR).

    Parameters:
    -----------
    path : str
        Path to the INI file.
    """
    with open(path, 'rb') as f:
        raw = f.read()
    try:
        raw.decode('utf-8')
        return 'utf-8'
    except UnicodeDecodeError:
        return 'cp949'

def write_sql_header(outfile, ini_file, spool_kind=None):
    """
    Write the tbsql script header.

    Parameters:
    -----------
    outfile : file object
        Output SQL file handle.
    ini_file : str
        Input INI file path (recorded as metadata in the spool output).

    Description:
    ------------
    - Writes the tbsql setup header for TPR data collection.
    - Declares bind variables, date range defaults, snapshot ID / snap_gid
      resolution, and HTML spool settings.

    Tibero snapshot semantics (verified on Tibero 7.2 TAC):
    - sys._tpr_* tables store per-interval deltas; the first snapshot of the
      selected window must be included (no Oracle-style baseline snap_id-1).
    - snap_id is an independent sequence per instance (thread# is 0-based), so a
      NODE report filters rows by thread# = :inst_no.
    - TAC reports must group instances by snap_gid, not snap_id; :gid_fr/:gid_to
      bound the global snapshot range for 'TAC' INIs.
    """

    # Distinguish Single-instance INIs ('NODE') from TAC INIs ('TAC': 2node..8node),
    # and embed a tag derived from the INI filename (without 'TprOne_') to identify the source INI.
    # To keep spool filenames concise, strip '_Single' for NODE INIs (redundant with 'NODE')
    # and compress '_with_ash' to '_ash'.
    # Examples:
    #   Simple_TprOne_2node.ini         -> TPR_TAC_Simple_2node_<host>_<inst>_<yyyymmdd_hhmiss>.html
    #   Simple_TprOne_Single.ini        -> TPR_NODE_Simple_<host>_<inst>_<yyyymmdd_hhmiss>.html
    #   Simple_TprOne_Single_with_ash.ini -> TPR_NODE_Simple_ash_<host>_<inst>_<yyyymmdd_hhmiss>.html
    ini_base = os.path.basename(str(ini_file))
    if spool_kind is None:
        # Guess from the filename convention ('single' => NODE, else TAC).
        spool_kind = 'NODE' if 'single' in ini_base.lower() else 'TAC'
    spool_tag = re.sub(r'\.ini$', '', ini_base, flags=re.IGNORECASE)
    spool_tag = spool_tag.replace('TprOne_', '')
    spool_tag = re.sub(r'[^A-Za-z0-9]+', '_', spool_tag).strip('_')
    spool_tag = spool_tag.replace('_with_ash', '_ash')
    if spool_kind == 'NODE':
        spool_tag = spool_tag.replace('_Single', '')

    if spool_kind == 'TAC':
        # TAC: cluster-wide snapshot grouping is by snap_gid because each
        # instance owns an independent snap_id sequence. The global range is
        # resolved into :gid_fr/:gid_to; :snap_fr/:snap_to still hold the local
        # instance range for sections that only need this node.
        snap_block = """declare
  v_min_snap number;
begin
  select min(snap_gid), max(snap_gid)
    into :gid_fr, :gid_to
    from sys._tpr_snapshot
   where snap_gid is not null
     and begin_interval_time >= to_date(:snapdate_from, 'YYYYMMDDHH24MI')
     and end_interval_time   <= to_date(:snapdate_to,   'YYYYMMDDHH24MI');
  if :gid_fr is null or :gid_to is null then
    raise_application_error(-20001, 'No TPR global snapshots found in requested snapdate_from..snapdate_to window.');
  end if;
  select nvl(min(snap_id), -1), nvl(max(snap_id), -1)
    into :snap_fr, :snap_to
    from sys._tpr_snapshot
   where "THREAD#" = :inst_no
     and snap_gid between :gid_fr and :gid_to;
end;
/"""
    else:
        # NODE: restrict to this instance's own snap_id sequence. TPR rows are
        # already interval deltas, so the first snapshot inside the window is
        # included directly and no baseline (snap_fr - 1) is subtracted.
        snap_block = """declare
  v_min_snap number;
  v_max_snap number;
begin
  select min(snap_id), max(snap_id)
    into v_min_snap, v_max_snap
    from sys._tpr_snapshot
   where "THREAD#" = :inst_no
     and begin_interval_time >= to_date(:snapdate_from, 'YYYYMMDDHH24MI')
     and end_interval_time   <= to_date(:snapdate_to,   'YYYYMMDDHH24MI');
  if v_min_snap is null or v_max_snap is null then
    raise_application_error(-20001, 'No TPR snapshots found in requested snapdate_from..snapdate_to window.');
  end if;
  :snap_fr := v_min_snap;
  :snap_to := v_max_snap;
  select nvl(min(snap_gid), -1), nvl(max(snap_gid), -1)
    into :gid_fr, :gid_to
    from sys._tpr_snapshot
   where "THREAD#" = :inst_no
     and snap_id between :snap_fr and :snap_to;
end;
/"""

    # Snapshot date window and bind variable setup header.
    # define stays ON: '&report_name' substitution requires it. Never put '&'
    # inside query string literals - tbsql would prompt for substitution.
    header_common = """Prompt ############### Define Date : YYYYMMDDHH24MI #################################
var snapdate_from number
exec :snapdate_from:=202601010000

var snapdate_to number
exec :snapdate_to:=202601150000
Prompt ##############################################################################

var inst_no number
var snap_fr number
var snap_to number
var gid_fr number
var gid_to number
var inst_no1 number
var inst_no2 number
var inst_no3 number
var inst_no4 number
var inst_no5 number
var inst_no6 number
var inst_no7 number
var inst_no8 number
var dbid number

exec select instance_number into :inst_no from v$instance;
exec select dbid into :dbid from v$database;
whenever sqlerror exit failure
""" + snap_block + """
whenever sqlerror continue
exec :inst_no1:=0
exec :inst_no2:=1
exec :inst_no3:=2
exec :inst_no4:=3
exec :inst_no5:=4
exec :inst_no6:=5
exec :inst_no7:=6
exec :inst_no8:=7

col dbid format 99999999999999999999
select :snap_fr snap_fr, :snap_to snap_to, :gid_fr gid_fr, :gid_to gid_to,
       :dbid dbid, :inst_no inst_no from dual;

set markup html on
set linesize 32767 trimspool on
set pages 50000
set feedback off
set termout off
set echo off"""

    # Spool filename header
    header_specific = """
prompt SQL>
set linesize 32767 pagesize 50000 trimspool on
set termout off time off
set sqlprompt "SQL> "
alter session set nls_date_format='YYYY-MM-DD HH24:MI:SS';
column report_name   new_value report_name   format a100;
select 'TPR_""" + spool_kind + '_' + spool_tag + """_' || host_name || '_' || instance_name|| '_'|| to_char(sysdate, 'yyyymmdd_hh24miss') || '.html' report_name from v$instance;

set echo on
set markup html on

spool &report_name
rem INI_NAME=""" + ini_file

    # Write complete header to output file
    outfile.write(header_common + header_specific + '\n')

def write_sql_footer(outfile):
    """
    Write the tbsql script footer.

    Parameters:
    -----------
    outfile : file object
        Output SQL file handle.

    Description:
    ------------
    Appends the 'exit;' command to terminate the tbsql session.
    """
    outfile.write("exit;\n")

def write_deferred_spools(outfile, deferred, ini_file):
    """
    Write SQL blocks assigned to separate spool files via 'SPOOL=<suffix>'.

    Closes the main spool file (TPR_xxx.html) with 'spool off', opens a new spool file
    named 'TPR_xxx_<suffix>.html' for each suffix (e.g. SPOOL=ash -> TPR_xxx_ASH.html),
    and writes the deferred sections. Includes 'rem INI_NAME=' in each new spool file
    so html_to_excel can locate the source INI configuration.

    Parameters:
    -----------
    outfile : file object
        Output SQL file handle.
    deferred : OrderedDict
        Mapping of {suffix: [section_sql_text, ...]}.
    ini_file : str
        Input INI file path.
    """
    for suffix, blocks in deferred.items():
        var = "spool_file_" + suffix.lower()
        outfile.write("spool off\n")
        outfile.write(f"column {var} new_value {var}\n")
        outfile.write(f"select replace('&report_name', '.html', '_{suffix}.html') {var} from dual;\n")
        outfile.write(f"spool &{var}\n")
        outfile.write(f"rem INI_NAME={ini_file}\n")
        for block in blocks:
            outfile.write(block)

def write_section_sql(outfile, section, sql_lines):
    """
    Write a single section's SQL query block to the output file.

    Parameters:
    -----------
    outfile : file object
        Output SQL file handle.
    section : str
        Current section name (from '[section]' in the INI file).
    sql_lines : list
        Lines collected from the section body.

    Description:
    ------------
    - Emits output only when the block starts with a SELECT statement.
    - Prepends a 'rem [section]' marker comment.
    - Appends a single terminating semicolon (';') on its own line.
    """
    if sql_lines:
        # Strip trailing semicolons and blank lines because a terminating ';' line
        # is appended explicitly below (avoiding an empty ';' line in tbsql).
        sql_lines = [ln.rstrip('\r\n') for ln in sql_lines]
        while sql_lines and not re.sub(r'[;\s]+$', '', sql_lines[-1]):
            sql_lines.pop()
        if not sql_lines:
            return
        sql_lines[-1] = re.sub(r'[;\s]+$', '', sql_lines[-1])
        first_line = sql_lines[0]

        # Check whether the first line contains 'select' (case-insensitive)
        if re.search(r'select', first_line, re.IGNORECASE):
            # Write section marker comment
            outfile.write(f"rem [{section}]\n")
            # Write first line
            outfile.write(first_line + "\n")
            # Write remaining lines
            if len(sql_lines) > 1:
                outfile.write("\n".join(sql_lines[1:]) + "\n")
            # Append SQL statement terminating semicolon
            outfile.write(";\n")

def convert_ini_to_sql(infile_name, outfile_name, kind=None):
    """
    Parse an INI file and generate the corresponding SQL collection script.

    Parameters:
    -----------
    infile_name : str
        Input INI file path.
    outfile_name : str
        Output SQL file path.
    kind : str, optional
        'NODE'/'TAC' override for the filename-based report-kind guess
        (CLI -k/--kind). When omitted, the guess is cross-checked against
        the INI body and a mismatch emits a stderr warning.

    Processing Steps:
    -----------------
    1. Read the INI file line by line.
    2. Ignore comment lines ('#'), disabled sections ('[#'), and blank lines.
    3. Detect '[section_name]' headers.
    4. Collect SQL query lines per section and route 'SPOOL=<suffix>' sections.
    5. Write the complete SQL script.

    INI File Structure:
    -------------------
    [section_name]
    SQL statement...

    [next_section]
    SQL statement...
    """
    # Temporary buffer for lines in the current section
    sql_lines = []
    # Name of the section currently being processed
    section = ""
    # Sections routed to separate spool files: {section_name: suffix} and buffered SQL blocks
    spool_map = {}
    deferred = OrderedDict()

    # Detect input encoding (UTF-8 first, falling back to cp949 for legacy INIs)
    with open(infile_name, 'r', encoding=detect_encoding(infile_name)) as ini_file:
        ini_lines = ini_file.readlines()

        # Resolve report kind: explicit -k/--kind wins; otherwise guess from the
        # filename and cross-check against bind references in the INI body.
        ini_base = os.path.basename(str(infile_name))
        guessed = 'NODE' if 'single' in ini_base.lower() else 'TAC'
        if kind:
            resolved_kind = kind
        else:
            resolved_kind = guessed
            body = ''.join(ini_lines)
            # TAC signature: multi-node binds (:inst_no1..N) or snap_gid used
            # beyond the single incidental '[dbinfo] Cluster ?' reference that
            # NODE INIs also contain.
            if (re.search(r':inst_no[1-9]', body)
                    or len(re.findall(r'snap_gid\b', body)) > 1):
                body_kind = 'TAC'
            elif re.search(r':inst_no\b', body):
                body_kind = 'NODE'
            else:
                body_kind = None
            if body_kind and body_kind != guessed:
                print(f"Warning: filename suggests {guessed} but INI references "
                      f"{'multi-node binds/snap_gid (TAC)' if body_kind == 'TAC' else ':inst_no only (NODE)'}. "
                      f"Use -k to override.", file=sys.stderr)

        # Open output SQL file in UTF-8
        with open(outfile_name, 'w', encoding='utf-8') as outfile:
            # Write SQL script header
            write_sql_header(outfile, infile_name, resolved_kind)

            def flush_section(sec, lines):
                """Emit section SQL, deferring sections configured with 'SPOOL=<suffix>'."""
                suffix = spool_map.get(sec)
                if not suffix:
                    write_section_sql(outfile, sec, lines)
                    return
                buf = io.StringIO()
                write_section_sql(buf, sec, lines)
                if buf.getvalue():
                    deferred.setdefault(suffix, []).append(buf.getvalue())

            # Process the INI file line by line
            for line in ini_lines:
                # Strip trailing LF (\n) and CR (\r)
                line = line.rstrip('\n\r')
                # Remove any remaining Windows CR (0x0D) characters
                line = line.replace('\x0D', '')

                # === Skip ignored lines ===

                # Ignore comment lines starting with '#'
                if line.startswith('#'):
                    continue
                # Ignore disabled section headers starting with '[#'
                if line.startswith('[#'):
                    continue
                # Ignore whitespace-only lines
                if re.match(r'^\s*$', line):
                    continue
                # Ignore empty lines
                if line == "":
                    continue

                # === Section header handling ===
                # Match '[section_name]' (\w+ matches letters, digits, and underscore)
                section_match = re.match(r'^\[(\w+)\]', line)
                if section_match:
                    # Flush buffered lines from the previous section
                    if sql_lines:
                        flush_section(section, sql_lines)
                        sql_lines = []  # Reset line buffer
                    # Record new section name
                    section = section_match.group(1)
                else:
                    # === SPOOL=<suffix>: route this section to a separate spool file ===
                    spool_match = re.match(r'^SPOOL\s*=\s*(\w+)\s*$', line, re.IGNORECASE)
                    if spool_match:
                        spool_map[section] = spool_match.group(1).upper()
                        continue
                    # === SQL/body line handling ===
                    # Treat non-header lines as candidate SQL body lines
                    sql_lines.append(line)
                    # Print buffered line when debug mode is enabled
                    if debug:
                        print(f"push ({line})")

            # Flush remaining lines at EOF
            if sql_lines:
                flush_section(section, sql_lines)

            # Emit deferred spool sections into separate files after closing the main spool
            write_deferred_spools(outfile, deferred, infile_name)

            # Write SQL script footer
            write_sql_footer(outfile)

def run_cli():
    """
    Command-line entry point.

    Steps:
    ------
    1. Parse command-line arguments.
    2. Verify that the input INI file exists.
    3. Determine the output SQL filename.
    4. Convert the INI file into a SQL script.
    5. Print completion or error status.

    Usage:
    ------
    python TprONE_ini_to_sql.py -i input.ini [-o output.sql]

    Options:
    --------
    -i, --input  : Required. Input INI file path.
    -o, --output : Optional. Output SQL file path (defaults to replacing .ini with .sql).
    """

    # Create ArgumentParser for CLI options
    parser = argparse.ArgumentParser(
        description='Convert INI file to SQL file',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Example:
  %(prog)s -i inifile.ini
  %(prog)s -i inifile.ini -o output.sql
        """
    )

    # Define required -i/--input argument
    parser.add_argument('-i', '--input',
                        required=True,
                        help='Configuration Review ini files')

    # Define optional -o/--output argument
    parser.add_argument('-o', '--output',
                        help='specify extracted file name (default: replace ini extension with .sql)')

    # Define optional -k/--kind argument (overrides filename-based detection)
    parser.add_argument('-k', '--kind',
                        choices=['NODE', 'TAC'],
                        help='override report kind detection (default: guess from ini filename)')

    # Parse command-line arguments
    args = parser.parse_args()

    # Determine output SQL filename
    if args.output:
        # Use explicit path from -o/--output
        outfile = args.output
    else:
        # Replace input extension with .sql (e.g. Simple_TprOne_Single.ini -> Simple_TprOne_Single.sql)
        base_name = os.path.splitext(args.input)[0]
        outfile = base_name + '.sql'

    # Check that the input INI file exists
    if not os.path.exists(args.input):
        print(f"Error: Input file '{args.input}' not found")
        sys.exit(1)

    # Parse INI file and write SQL script
    try:
        convert_ini_to_sql(args.input, outfile, kind=args.kind)
        # Print success message
        print(f"Successfully converted '{args.input}' to '{outfile}'")
    except Exception as e:
        # Print error message and exit with non-zero status
        print(f"Error processing file: {e}")
        sys.exit(1)

# Invoke run_cli() only when executed directly (not when imported as a module)
if __name__ == "__main__":
    run_cli()
