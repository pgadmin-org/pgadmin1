Attribute VB_Name = "basSQL"
' pgAdmin - PostgreSQL db Administration/Management for Win32
' Copyright (C) 1998 - 2001, Dave Page

' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.

' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.

' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

Option Explicit
Dim SQL_PGADMIN_SELECT_CLAUSE_FUNCTIONS As String
Dim SQL_PGADMIN_SELECT_CLAUSE_VIEWS As String
Dim SQL_PGADMIN_SELECT_CLAUSE_TRIGGERS As String

Public Sub Chk_HelperObjects()
On Error GoTo Err_Handler

Dim rsParam As New Recordset
Dim SQL_PGADMIN_SYS As String
Dim SQL_PGADMIN_PARAM As String
Dim SQL_INS_PGADMIN_PARAM1 As String
Dim SQL_INS_PGADMIN_PARAM2 As String
Dim SQL_INS_PGADMIN_PARAM3 As String
Dim SQL_INS_PGADMIN_PARAM4 As String
Dim SQL_PGADMIN_LOG As String
Dim SQL_PGADMIN_SEQ_CACHE As String
Dim SQL_PGADMIN_TABLE_CACHE As String
Dim SQL_PGADMIN_GET_DESC As String
Dim SQL_PGADMIN_GET_COL_DEF As String
Dim SQL_PGADMIN_GET_HANDLER As String
Dim SQL_PGADMIN_GET_TYPE As String
Dim SQL_PGADMIN_GET_ROWS As String
Dim SQL_PGADMIN_GET_SEQUENCE As String
Dim SQL_PGADMIN_GET_FUNCTION_NAME As String
Dim SQL_PGADMIN_GET_FUNCTION_ARGUMENTS As String
Dim SQL_PGADMIN_CHECKS As String
Dim SQL_PGADMIN_DATABASES As String
Dim SQL_PGADMIN_FUNCTIONS As String
Dim SQL_PGADMIN_GROUPS As String
Dim SQL_PGADMIN_INDEXES As String
Dim SQL_PGADMIN_LANGUAGES As String
Dim SQL_PGADMIN_SEQUENCES As String
Dim SQL_PGADMIN_TABLES As String
Dim SQL_PGADMIN_USERS As String
Dim SQL_PGADMIN_TRIGGERS As String
Dim SQL_PGADMIN_VIEWS As String
Dim SQL_PGADMIN_DEV_FUNCTIONS As String
Dim SQL_PGADMIN_DEV_TRIGGERS As String
Dim SQL_PGADMIN_DEV_VIEWS As String
Dim SQL_PGADMIN_DEV_DEPENDENCIES As String

SQL_PGADMIN_PARAM = "CREATE TABLE pgadmin_param(param_id int4, param_value text, param_desc text)"
SQL_INS_PGADMIN_PARAM1 = "INSERT INTO pgadmin_param VALUES ('1', '" & Str(SSO_VERSION) & "', 'SSO Version')"
SQL_INS_PGADMIN_PARAM2 = "INSERT INTO pgadmin_param VALUES (2, 'N', 'Revision Tracking enabled?')"
SQL_INS_PGADMIN_PARAM3 = "INSERT INTO pgadmin_param VALUES (3, '1.0', 'Revision Tracking version')"
SQL_INS_PGADMIN_PARAM4 = "INSERT INTO pgadmin_param VALUES (4, 'N', 'Development Mode?')"
SQL_PGADMIN_LOG = "CREATE TABLE pgadmin_rev_log(event_timestamp timestamp DEFAULT now(), username text, version float4, query text)"
SQL_PGADMIN_SEQ_CACHE = "CREATE TABLE pgadmin_seq_cache(sequence_oid oid, sequence_last_value int4, sequence_increment_by int4, sequence_max_value int4, sequence_min_value int4, sequence_cache_value int4, sequence_is_cycled text, sequence_timestamp timestamp DEFAULT now())"
SQL_PGADMIN_TABLE_CACHE = "CREATE TABLE pgadmin_table_cache(table_oid oid, table_rows int4, table_timestamp timestamp DEFAULT now())"
SQL_PGADMIN_GET_DESC = "CREATE FUNCTION pgadmin_get_desc(oid) RETURNS text AS 'SELECT description FROM pg_description WHERE objoid = $1' LANGUAGE 'sql'"
SQL_PGADMIN_GET_COL_DEF = "CREATE FUNCTION pgadmin_get_col_def(oid, int4) RETURNS text AS 'SELECT adsrc FROM pg_attrdef WHERE adrelid = $1 AND adnum = $2' LANGUAGE 'sql'"
SQL_PGADMIN_GET_HANDLER = "CREATE FUNCTION pgadmin_get_handler(oid) RETURNS text AS 'SELECT proname::text FROM pg_proc WHERE oid = $1' LANGUAGE 'sql'"
SQL_PGADMIN_GET_TYPE = "CREATE FUNCTION pgadmin_get_type(oid) RETURNS text AS 'SELECT typname::text FROM pg_type WHERE oid = $1' LANGUAGE 'sql'"
SQL_PGADMIN_GET_ROWS = "CREATE FUNCTION pgadmin_get_rows(oid) RETURNS pgadmin_table_cache AS 'SELECT DISTINCT ON(table_oid) * FROM pgadmin_table_cache WHERE table_oid = $1 ORDER BY table_oid, table_timestamp DESC' LANGUAGE 'sql'"
SQL_PGADMIN_GET_SEQUENCE = "CREATE FUNCTION pgadmin_get_sequence(oid) RETURNS pgadmin_seq_cache AS 'SELECT DISTINCT ON(sequence_oid) * FROM pgadmin_seq_cache WHERE sequence_oid = $1 ORDER BY sequence_oid, sequence_timestamp DESC' LANGUAGE 'sql'"
SQL_PGADMIN_GET_FUNCTION_NAME = "CREATE FUNCTION pgadmin_get_function_name(oid) RETURNS name AS 'SELECT function_name FROM pgadmin_functions WHERE function_oid = $1' LANGUAGE 'sql'"
SQL_PGADMIN_GET_FUNCTION_ARGUMENTS = "CREATE FUNCTION pgadmin_get_function_arguments(oid) RETURNS text AS 'SELECT function_arguments FROM pgadmin_functions WHERE function_oid = $1' LANGUAGE 'sql'"

SQL_PGADMIN_CHECKS = _
  "CREATE VIEW pgadmin_checks AS SELECT " & _
  "  r.oid AS check_oid, " & _
  "  r.rcname AS check_name, " & _
  "  c.oid AS check_table_oid, " & _
  "  c.relname AS check_table_name, " & _
  "  r.rcsrc AS check_definition, " & _
  "  pgadmin_get_desc(r.oid) AS check_comments " & _
  "FROM pg_relcheck r, pg_class c " & _
  "WHERE r.rcrelid = c.oid"
  
SQL_PGADMIN_DATABASES = _
  "CREATE VIEW pgadmin_databases AS SELECT " & _
  "  d.oid AS database_oid, " & _
  "  d.datname AS database_name, " & _
  "  d.datpath AS database_path, " & _
  "  pg_get_userbyid(d.datdba) AS database_owner, " & _
  "  pgadmin_get_desc(d.oid) AS database_comments " & _
  "FROM pg_database d"
  
SQL_PGADMIN_FUNCTIONS = _
  "CREATE VIEW pgadmin_functions AS SELECT " & _
  "  p.oid AS function_oid, p.proname AS function_name, pg_get_userbyid(p.proowner) AS function_owner, rtrim(trim(CASE WHEN (pgadmin_get_type(p.proargtypes[0]) NOTNULL) THEN pgadmin_get_type(p.proargtypes[0]) || ', ' ELSE '' END || CASE WHEN (pgadmin_get_type(p.proargtypes[1]) NOTNULL) THEN pgadmin_get_type(p.proargtypes[1]) || ', ' ELSE '' END || CASE WHEN (pgadmin_get_type(p.proargtypes[2]) NOTNULL) THEN pgadmin_get_type(p.proargtypes[2]) || ', ' ELSE '' END || CASE WHEN (pgadmin_get_type(p.proargtypes[3]) NOTNULL) THEN pgadmin_get_type(p.proargtypes[3]) || ', ' ELSE '' END || CASE WHEN (pgadmin_get_type(p.proargtypes[4]) NOTNULL) THEN pgadmin_get_type(p.proargtypes[4]) || ', ' ELSE '' END || " & _
  "  CASE WHEN (pgadmin_get_type(p.proargtypes[5]) NOTNULL) THEN pgadmin_get_type(p.proargtypes[5]) || ', ' ELSE '' END || CASE WHEN (pgadmin_get_type(p.proargtypes[6]) NOTNULL) THEN pgadmin_get_type(p.proargtypes[6]) || ', ' ELSE '' END || CASE WHEN (pgadmin_get_type(p.proargtypes[7]) NOTNULL) THEN pgadmin_get_type(p.proargtypes[7]) || ', ' ELSE '' END || CASE WHEN (pgadmin_get_type(p.proargtypes[8]) NOTNULL) THEN pgadmin_get_type(p.proargtypes[8]) || ', ' ELSE '' END || CASE WHEN (pgadmin_get_type(p.proargtypes[9]) NOTNULL) THEN pgadmin_get_type(p.proargtypes[9]) || ', ' ELSE '' END || CASE WHEN (pgadmin_get_type(p.proargtypes[10]) NOTNULL) THEN pgadmin_get_type(p.proargtypes[10]) || ', ' ELSE '' END || " & _
  "  CASE WHEN (pgadmin_get_type(p.proargtypes[11]) NOTNULL) THEN pgadmin_get_type(p.proargtypes[11]) || ', ' ELSE '' END || CASE WHEN (pgadmin_get_type(p.proargtypes[12]) NOTNULL) THEN pgadmin_get_type(p.proargtypes[12]) || ', ' ELSE '' END || CASE WHEN (pgadmin_get_type(p.proargtypes[13]) NOTNULL) THEN pgadmin_get_type(p.proargtypes[13]) || ', ' ELSE '' END || CASE WHEN (pgadmin_get_type(p.proargtypes[14]) NOTNULL) THEN pgadmin_get_type(p.proargtypes[14]) || ', ' ELSE '' END || CASE WHEN (pgadmin_get_type(p.proargtypes[15]) NOTNULL) THEN pgadmin_get_type(p.proargtypes[15]) || ', ' ELSE '' END), ',') AS function_arguments, " & _
  "  pgadmin_get_type(p.prorettype) AS function_returns, p.prosrc AS function_source, l.lanname AS function_language, " & _
  "  pgadmin_get_desc(p.oid) AS function_comments " & _
  "FROM pg_proc p, pg_language l " & _
  "WHERE p.prolang = l.oid " & _
  "AND p.proname NOT LIKE '%_call_handler'"
   
SQL_PGADMIN_GROUPS = _
  "CREATE VIEW pgadmin_groups AS SELECT " & _
  "oid AS group_oid, groname AS group_name, grosysid AS group_id, grolist As group_members " & _
  "FROM pg_group"

SQL_PGADMIN_INDEXES = _
  "CREATE VIEW pgadmin_indexes AS SELECT " & _
  "  i.oid AS index_oid, i.relname AS index_name, c.relname AS index_table, pg_get_userbyid(i.relowner) AS index_owner, " & _
  "  CASE WHEN x.indislossy = TRUE THEN 'Yes'::text ELSE 'No'::text END AS index_is_lossy, CASE WHEN x.indisunique = TRUE THEN 'Yes'::text ELSE 'No'::text END AS index_is_unique, CASE WHEN x.indisprimary = TRUE THEN 'Yes'::text ELSE 'No'::text END AS index_is_primary, " & _
  "  pgadmin_get_desc(i.oid) AS index_comments, " & _
  "  pg_get_indexdef(x.indexrelid) AS index_definition, a.oid AS column_oid, a.attname AS column_name,  a.attnum AS column_position,  t.typname As column_type, " & _
  "  CASE WHEN ((a.attlen = -1) AND ((a.atttypmod)::int4 = (-1)::int4)) THEN (0)::int4 ELSE CASE WHEN a.attlen = -1 THEN CASE WHEN ((t.typname = 'bpchar') OR (t.typname = 'char') OR (t.typname = 'varchar')) THEN (a.atttypmod -4)::int4 ELSE (a.atttypmod)::int4 END ELSE (a.attlen)::int4 END END AS column_length, " & _
  "  pgadmin_get_desc(a.oid) AS column_comments " & _
  "FROM pg_index x, pg_attribute a, pg_type t, pg_class c,pg_class i " & _
  "WHERE a.atttypid = t.oid AND a.attrelid = i.oid AND ((c.oid = x.indrelid) AND (i.oid = x.indexrelid))"

SQL_PGADMIN_LANGUAGES = _
  "CREATE VIEW pgadmin_languages AS SELECT " & _
  "  l.oid AS language_oid, " & _
  "  l.lanname AS language_name, " & _
  "  l.lancompiler AS language_compiler, " & _
  "  CASE WHEN l.lanpltrusted = TRUE THEN 'Yes'::text ELSE 'No'::text END AS language_is_trusted, " & _
  "  pgadmin_get_handler(lanplcallfoid) AS language_handler, " & _
  "  pgadmin_get_desc(l.oid) AS language_comments " & _
  "FROM pg_language l"
  
SQL_PGADMIN_SEQUENCES = _
  "CREATE VIEW pgadmin_sequences AS SELECT  " & _
  "  c.oid AS sequence_oid, c.relname AS sequence_name, pg_get_userbyid(c.relowner) AS sequence_owner, c.relacl AS sequence_acl, sequence_last_value(pgadmin_get_sequence(c.oid)) AS sequence_last_value, " & _
  "  sequence_increment_by(pgadmin_get_sequence(c.oid)) AS sequence_increment_by, sequence_max_value(pgadmin_get_sequence(c.oid)) AS sequence_max_value, sequence_min_value(pgadmin_get_sequence(c.oid)) AS sequence_min_value, " & _
  "  sequence_cache_value(pgadmin_get_sequence(c.oid)) AS sequence_cache_value, sequence_is_cycled(pgadmin_get_sequence(c.oid)) AS sequence_is_cycled, " & _
  "  pgadmin_get_desc(c.oid) AS sequence_comments " & _
  "FROM pg_class c " & _
  "WHERE c.relkind = 'S'"

SQL_PGADMIN_TABLES = _
  "CREATE VIEW pgadmin_tables AS SELECT " & _
  "  c.oid AS table_oid, c.relname AS table_name, pg_get_userbyid(c.relowner) AS table_owner, c.relacl AS table_acl, " & _
  "  CASE WHEN c.relhasindex = TRUE THEN 'Yes'::text ELSE 'No'::text END AS table_has_indexes, CASE WHEN c.relhasrules = TRUE THEN 'Yes'::text ELSE 'No'::text END AS table_has_rules, CASE WHEN c.relisshared = TRUE THEN 'Yes'::text ELSE 'No'::text END  AS table_is_shared, CASE WHEN c.relhaspkey = TRUE THEN 'Yes'::text ELSE 'No'::text END AS table_has_primarykey, CASE WHEN c.reltriggers > 0 THEN 'Yes'::text ELSE 'No'::text END AS table_has_triggers, " & _
  "  table_rows(pgadmin_get_rows(c.oid)) AS table_rows, pgadmin_get_desc(c.oid) AS table_comments, a.oid AS column_oid, a.attname AS column_name, a.attnum AS column_position, t.typname As column_type,  " & _
  "  CASE WHEN ((a.attlen = -1) AND ((a.atttypmod)::int4 = (-1)::int4)) THEN (0)::int4 ELSE CASE WHEN a.attlen = -1 THEN " & _
  "  CASE WHEN ((t.typname = 'bpchar') OR (t.typname = 'char') OR (t.typname = 'varchar')) THEN (a.atttypmod -4)::int4 ELSE (a.atttypmod)::int4 END " & _
  "  ELSE (a.attlen)::int4 END END AS column_length, " & _
  "  CASE WHEN a.attnotnull = TRUE THEN 'Yes'::text ELSE 'No'::text END AS column_not_null, CASE WHEN a.atthasdef = TRUE THEN 'Yes'::text ELSE 'No'::text END AS column_has_default,  " & _
  "  CASE WHEN (pgadmin_get_col_def(c.oid, a.attnum) NOTNULL) THEN pgadmin_get_col_def(c.oid, a.attnum) ELSE '' END AS column_default, pgadmin_get_desc(a.oid) AS column_comments " & _
  "FROM pg_attribute a, pg_type t, pg_class c " & _
  "WHERE a.atttypid = t.oid AND a.attrelid = c.oid AND (((c.relkind::char = 'r'::char) OR (c.relkind::char = 's'::char)) AND (NOT (EXISTS (SELECT pg_rewrite.rulename FROM pg_rewrite WHERE ((pg_rewrite.ev_class = c.oid) AND (pg_rewrite.ev_type::char = '1'::char))))))"
  
SQL_PGADMIN_USERS = _
  "CREATE VIEW pgadmin_users AS SELECT " & _
  "  oid AS user_oid, usename AS user_name, usesysid AS user_id, " & _
  "  CASE WHEN usecreatedb = TRUE THEN 'Yes'::text ELSE 'No'::text END AS user_create_dbs, " & _
  "  CASE WHEN usesuper = TRUE THEN 'Yes'::text ELSE 'No'::text END AS user_superuser, " & _
  "  valuntil As user_expires " & _
  "FROM pg_shadow"

SQL_PGADMIN_TRIGGERS = _
  "CREATE VIEW pgadmin_triggers AS SELECT " & _
  " t.oid AS trigger_oid, t.tgname AS trigger_name, c.relname AS trigger_table, " & _
  " pgadmin_get_function_name (tgfoid) AS trigger_function, pgadmin_get_function_arguments (tgfoid) AS trigger_arguments, t.tgtype AS trigger_type, " & _
  " pgadmin_get_desc(t.oid) AS trigger_comments " & _
  "FROM pg_trigger t, pg_class c " & _
  "WHERE c.oid = t.tgrelid"

SQL_PGADMIN_VIEWS = _
  "CREATE VIEW pgadmin_views AS SELECT " & _
  "  c.oid AS view_oid, " & _
  "  c.relname AS view_name, " & _
  "  pg_get_userbyid(c.relowner) AS view_owner, " & _
  "  c.relacl AS view_acl, " & _
  "  pgadmin_get_desc(c.oid) AS view_comments " & _
  "FROM " & _
  "  pg_class c " & _
  "WHERE " & _
  "  ((c.relhasrules AND (EXISTS (SELECT r.rulename FROM pg_rewrite r WHERE ((r.ev_class = c.oid) AND (r.ev_type::char = '1'::char))))) OR c.relkind = 'v')"

' SQL_PGADMIN_SELECT_CLAUSE_FUNCTIONS is used later in cmp_Project_Move_Functions
SQL_PGADMIN_SELECT_CLAUSE_FUNCTIONS = "function_name NOT LIKE '%_call_handler' " & _
  "  AND function_name NOT LIKE 'pgadmin_%' " & _
  "  AND function_name NOT LIKE 'pg_%' " & _
  "  AND function_oid > " & LAST_SYSTEM_OID & _
  "  ORDER BY function_name ;"

SQL_PGADMIN_DEV_FUNCTIONS = "CREATE TABLE pgadmin_dev_functions AS SELECT * FROM pgadmin_functions WHERE " & SQL_PGADMIN_SELECT_CLAUSE_FUNCTIONS & _
  "  ALTER TABLE pgadmin_dev_functions ADD function_iscompiled boolean DEFAULT 'f'  ;" & _
  "  TRUNCATE TABLE pgadmin_dev_functions ;"

' SQL_PGADMIN_SELECT_CLAUSE_TRIGGERS is used later in cmp_Project_Move_Triggers
SQL_PGADMIN_SELECT_CLAUSE_TRIGGERS = "trigger_oid > " & LAST_SYSTEM_OID & _
  "  AND trigger_name NOT LIKE 'pgadmin_%' " & _
  "  AND trigger_name NOT LIKE 'pg_%' " & _
  "  AND trigger_name NOT LIKE 'RI_ConstraintTrigger_%' " & _
  "  ORDER BY trigger_name; "

SQL_PGADMIN_DEV_TRIGGERS = "CREATE TABLE pgadmin_dev_triggers AS SELECT * FROM pgadmin_triggers WHERE " & SQL_PGADMIN_SELECT_CLAUSE_TRIGGERS & _
  "  ALTER TABLE pgadmin_dev_triggers ADD trigger_iscompiled boolean DEFAULT 'f'  ;" & _
  "  TRUNCATE TABLE pgadmin_dev_triggers ;"

' SQL_PGADMIN_SELECT_CLAUSE_VIEWS is used later in cmp_Project_Move_Views
SQL_PGADMIN_SELECT_CLAUSE_VIEWS = "view_oid > " & LAST_SYSTEM_OID & _
  "  AND view_name NOT LIKE 'pgadmin_%' " & _
  "  AND view_name NOT LIKE 'pg_%' " & _
  "  ORDER BY view_name; "

SQL_PGADMIN_DEV_VIEWS = "CREATE TABLE pgadmin_dev_views AS SELECT " & _
  "  view_oid, view_name, view_owner, view_comments " & _
  "  FROM pgadmin_views WHERE " & SQL_PGADMIN_SELECT_CLAUSE_VIEWS & _
  "  ALTER TABLE pgadmin_dev_views ADD view_definition text;  " & _
  "  ALTER TABLE pgadmin_dev_views ADD view_acl text;  " & _
  "  ALTER TABLE pgadmin_dev_views ADD view_iscompiled boolean DEFAULT 'f'  ;" & _
  "  TRUNCATE TABLE pgadmin_dev_views ;"

SQL_PGADMIN_DEV_DEPENDENCIES = "CREATE TABLE pgadmin_dev_dependencies (" & _
  " dependency_project_oid int4," & _
  " dependency_parent_object text," & _
  " dependency_parent_name   text," & _
  " dependency_child_object  text," & _
  " dependency_child_name    text);"

  'If the SSO Version on the server doesn't exist or is lower than
  'that defined in SSO_VERSION then drop all SSO's. If pgadmin_param
  'doesn't exist then we'll get a new set of SSO's anyway.
  
  If ObjectExists("pgadmin_param", tTable) > 0 Then
    StartMsg "Checking SSO Version..."
    If rsParam.State <> adStateClosed Then rsParam.Close
    LogMsg "Executing: SELECT param_value FROM pgadmin_param WHERE param_id = 1"
    rsParam.Open "SELECT param_value FROM pgadmin_param WHERE param_id = 1", gConnection, adOpenForwardOnly
    If Not rsParam.EOF Then 'Param 1 exists so check it.
      If Val(rsParam!param_value) < SSO_VERSION Then Drop_Objects False
    Else 'Param 1 doesn't exist so drop SSO's to be safe.
      If Not SuperuserChk Then Exit Sub
      Drop_Objects False
    End If
    EndMsg
  End If
  
  'Check Descriptions table. If it exist then migrate to pg_description & drop it.
  If ObjectExists("pgadmin_desc", tTable) <> 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Removing old pgAdmin Descriptions Table..."
    LogMsg "Executing: INSERT INTO pg_description SELECT * FROM pgadmin_desc"
    gConnection.Execute "INSERT INTO pg_description SELECT * FROM pgadmin_desc"
    LogMsg "Executing: DROP TABLE pgadmin_desc"
    gConnection.Execute "DROP TABLE pgadmin_desc"
    EndMsg
  End If
  
  'Drop any old pgadmin_get_pgdesc functions
  If ObjectExists("pgadmin_get_pgdesc", tFunction) <> 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Dropping old pgAdmin Description Lookup Function..."
    LogMsg "Executing: DROP FUNCTION pgadmin_get_pgdesc(oid)"
    gConnection.Execute "DROP FUNCTION pgadmin_get_pgdesc(oid)"
    EndMsg
  End If
  
  If ObjectExists("pgadmin_param", tTable) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Parameter Table..."
    LogMsg "Executing: " & SQL_PGADMIN_PARAM
    gConnection.Execute SQL_PGADMIN_PARAM
    LogMsg "Executing: " & SQL_INS_PGADMIN_PARAM1
    gConnection.Execute SQL_INS_PGADMIN_PARAM1
    LogMsg "Executing: GRANT all ON pgadmin_param TO public"
    gConnection.Execute "GRANT all ON pgadmin_param TO public"
    If ObjectExists("pgadmin_sys", tTable) <> 0 Then
      If rsParam.State <> adStateClosed Then rsParam.Close
      LogMsg "Executing: SELECT * FROM pgadmin_sys WHERE id = 1"
      rsParam.Open "SELECT * FROM pgadmin_sys WHERE id = 1", gConnection
      If Not rsParam.EOF Then
        LogMsg "Executing: INSERT INTO pgadmin_param VALUES(2, '" & rsParam!Tracking & "', 'Revision Tracking enabled?')"
        gConnection.Execute "INSERT INTO pgadmin_param VALUES(2, '" & rsParam!Tracking & "', 'Revision Tracking enabled?')"
        LogMsg "Executing: INSERT INTO pgadmin_param VALUES(3, '" & rsParam!Version & "', 'Revision Tracking version')"
        gConnection.Execute "INSERT INTO pgadmin_param VALUES(3, '" & rsParam!Version & "', 'Revision Tracking version')"
        MsgBox "Your Revision Tracking settings have been migrated to the new format used by this version of pgAdmin. Any subsequent database changes made by earlier versions of pgAdmin will not be logged correctly.", vbInformation, "Warning"
      Else
        LogMsg "Executing: " & SQL_INS_PGADMIN_PARAM2
        gConnection.Execute SQL_INS_PGADMIN_PARAM2
        LogMsg "Executing: " & SQL_INS_PGADMIN_PARAM3
        gConnection.Execute SQL_INS_PGADMIN_PARAM3
        LogMsg "Executing: " & SQL_INS_PGADMIN_PARAM4
        gConnection.Execute SQL_INS_PGADMIN_PARAM4
      End If
      If rsParam.State <> adStateClosed Then rsParam.Close
      LogMsg "Executing: DROP TABLE pgadmin_sys"
      gConnection.Execute "DROP TABLE pgadmin_sys"
    End If
    EndMsg
  End If
  
  'Repair pgadmin_param if necessary
  If rsParam.State <> adStateClosed Then rsParam.Close
  LogMsg "Executing: SELECT * FROM pgadmin_param WHERE param_id = 1"
  rsParam.Open "SELECT * FROM pgadmin_param WHERE param_id = 1", gConnection, adOpenForwardOnly
  If rsParam.EOF Then
    LogMsg "Executing: " & SQL_INS_PGADMIN_PARAM1
    gConnection.Execute SQL_INS_PGADMIN_PARAM1
  End If
  If rsParam.State <> adStateClosed Then rsParam.Close
  LogMsg "Executing: SELECT * FROM pgadmin_param WHERE param_id = 2"
  rsParam.Open "SELECT * FROM pgadmin_param WHERE param_id = 2", gConnection, adOpenForwardOnly
  If rsParam.EOF Then
    LogMsg "Executing: " & SQL_INS_PGADMIN_PARAM2
    gConnection.Execute SQL_INS_PGADMIN_PARAM2
  End If
  If rsParam.State <> adStateClosed Then rsParam.Close
  LogMsg "Executing: SELECT * FROM pgadmin_param WHERE param_id = 3"
  rsParam.Open "SELECT * FROM pgadmin_param WHERE param_id = 3", gConnection, adOpenForwardOnly
  If rsParam.EOF Then
    LogMsg "Executing: " & SQL_INS_PGADMIN_PARAM3
    gConnection.Execute SQL_INS_PGADMIN_PARAM3
  End If
  If rsParam.State <> adStateClosed Then rsParam.Close
  LogMsg "Executing: SELECT * FROM pgadmin_param WHERE param_id = 4"
  rsParam.Open "SELECT * FROM pgadmin_param WHERE param_id = 4", gConnection, adOpenForwardOnly
  If rsParam.EOF Then
    LogMsg "Executing: " & SQL_INS_PGADMIN_PARAM4
    gConnection.Execute SQL_INS_PGADMIN_PARAM4
  End If
  
  If ObjectExists("pgadmin_rev_log", tTable) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Revision Log Table..."
    LogMsg "Executing: " & SQL_PGADMIN_LOG
    gConnection.Execute SQL_PGADMIN_LOG
    LogMsg "Executing: GRANT all ON pgadmin_rev_log TO public"
    gConnection.Execute "GRANT all ON pgadmin_rev_log TO public"
    If ObjectExists("pgadmin_log", tTable) <> 0 Then
      Dim rsLog As New Recordset
      rsLog.Open "SELECT * FROM pgadmin_log", gConnection
      While Not rsLog.EOF
        LogMsg "Executing: INSERT INTO pgadmin_rev_log (event_timestamp, username, version, query) VALUES ('" & rsLog!event_date & " 00:00:01', '" & rsLog!Username & "', '" & rsLog!Version & "', '" & rsLog!Query & "')"
        gConnection.Execute "INSERT INTO pgadmin_rev_log (event_timestamp, username, version, query) VALUES ('" & rsLog!event_date & " 00:00:01', '" & rsLog!Username & "', '" & rsLog!Version & "', '" & rsLog!Query & "')"
        rsLog.MoveNext
      Wend
      If rsLog.State <> adStateClosed Then rsLog.Close
      LogMsg "Executing: DROP TABLE pgadmin_log"
      gConnection.Execute "DROP TABLE pgadmin_log"
      MsgBox "Your Revision Tracking log has been migrated to the new format used by this version of pgAdmin. Any subsequent database changes made by earlier versions of pgAdmin will not be logged correctly.", vbInformation, "Warning"
      End If
    EndMsg
  End If
  If ObjectExists("pgadmin_seq_cache", tTable) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Sequence Cache Table..."
    LogMsg "Executing: " & SQL_PGADMIN_SEQ_CACHE
    gConnection.Execute SQL_PGADMIN_SEQ_CACHE
    LogMsg "Executing: GRANT all ON pgadmin_seq_cache TO public"
    gConnection.Execute "GRANT all ON pgadmin_seq_cache TO public"
    EndMsg
  End If
  If ObjectExists("pgadmin_table_cache", tTable) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Table Cache Table..."
    LogMsg "Executing: " & SQL_PGADMIN_TABLE_CACHE
    gConnection.Execute SQL_PGADMIN_TABLE_CACHE
    LogMsg "Executing: GRANT all ON pgadmin_table_cache TO public"
    gConnection.Execute "GRANT all ON pgadmin_table_cache TO public"
    EndMsg
  End If
  If ObjectExists("pgadmin_get_desc", tFunction) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Description Lookup Function..."
    LogMsg "Executing: " & SQL_PGADMIN_GET_DESC
    gConnection.Execute SQL_PGADMIN_GET_DESC
    EndMsg
  End If
  If ObjectExists("pgadmin_get_col_def", tFunction) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Column Default Lookup Function..."
    LogMsg "Executing: " & SQL_PGADMIN_GET_COL_DEF
    gConnection.Execute SQL_PGADMIN_GET_COL_DEF
    EndMsg
  End If
  If ObjectExists("pgadmin_get_handler", tFunction) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Language Handler Lookup Function..."
    LogMsg "Executing: " & SQL_PGADMIN_GET_HANDLER
    gConnection.Execute SQL_PGADMIN_GET_HANDLER
    EndMsg
  End If
  If ObjectExists("pgadmin_get_type", tFunction) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Type Lookup Function..."
    LogMsg "Executing: " & SQL_PGADMIN_GET_TYPE
    gConnection.Execute SQL_PGADMIN_GET_TYPE
    EndMsg
  End If
  If ObjectExists("pgadmin_get_rows", tFunction) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Row Count Lookup Function..."
    LogMsg "Executing: " & SQL_PGADMIN_GET_ROWS
    gConnection.Execute SQL_PGADMIN_GET_ROWS
    EndMsg
  End If
  If ObjectExists("pgadmin_get_sequence", tFunction) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Sequence Lookup Function..."
    LogMsg "Executing: " & SQL_PGADMIN_GET_SEQUENCE
    gConnection.Execute SQL_PGADMIN_GET_SEQUENCE
    EndMsg
  End If
  If ObjectExists("pgadmin_databases", tTable) <> 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Dropping corrupted pgAdmin Databases View..."
    LogMsg "Executing: DROP TABLE pgadmin_databases"
    gConnection.Execute "DROP TABLE pgadmin_databases"
    EndMsg
  End If
  If ObjectExists("pgadmin_databases", tView) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Databases View..."
    LogMsg "Executing: " & SQL_PGADMIN_DATABASES
    gConnection.Execute SQL_PGADMIN_DATABASES
    LogMsg "Executing: GRANT all ON pgadmin_databases TO public"
    gConnection.Execute "GRANT all ON pgadmin_databases TO public"
    EndMsg
  End If
  If ObjectExists("pgadmin_checks", tTable) <> 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Dropping corrupted pgAdmin Checks View..."
    LogMsg "Executing: DROP TABLE pgadmin_checks"
    gConnection.Execute "DROP TABLE pgadmin_checks"
    EndMsg
  End If
  If ObjectExists("pgadmin_checks", tView) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Checks View..."
    LogMsg "Executing: " & SQL_PGADMIN_CHECKS
    gConnection.Execute SQL_PGADMIN_CHECKS
    LogMsg "Executing: GRANT all ON pgadmin_checks TO public"
    gConnection.Execute "GRANT all ON pgadmin_checks TO public"
    EndMsg
  End If
  If ObjectExists("pgadmin_functions", tTable) <> 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Dropping corrupted pgAdmin Functions View..."
    LogMsg "Executing: DROP TABLE pgadmin_functions"
    gConnection.Execute "DROP TABLE pgadmin_functions"
    EndMsg
  End If
  If ObjectExists("pgadmin_functions", tView) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Functions View..."
    LogMsg "Executing: " & SQL_PGADMIN_FUNCTIONS
    gConnection.Execute SQL_PGADMIN_FUNCTIONS
    LogMsg "Executing: GRANT all ON pgadmin_functions TO public"
    gConnection.Execute "GRANT all ON pgadmin_functions TO public"
    EndMsg
  End If
  
  If ObjectExists("pgadmin_get_function_name", tFunction) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Function Name Lookup Function..."
    LogMsg "Executing: " & SQL_PGADMIN_GET_FUNCTION_NAME
    gConnection.Execute SQL_PGADMIN_GET_FUNCTION_NAME
    EndMsg
  End If
  If ObjectExists("pgadmin_get_function_arguments", tFunction) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Sequence Lookup Function..."
    LogMsg "Executing: " & SQL_PGADMIN_GET_FUNCTION_ARGUMENTS
    gConnection.Execute SQL_PGADMIN_GET_FUNCTION_ARGUMENTS
    EndMsg
  End If
  
  If ObjectExists("pgadmin_groups", tTable) <> 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Dropping corrupted pgAdmin Groups View..."
    LogMsg "Executing: DROP TABLE pgadmin_groups"
    gConnection.Execute "DROP TABLE pgadmin_groups"
    EndMsg
  End If
  If ObjectExists("pgadmin_groups", tView) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Groups View..."
    LogMsg "Executing: " & SQL_PGADMIN_GROUPS
    gConnection.Execute SQL_PGADMIN_GROUPS
    LogMsg "Executing: GRANT all ON pgadmin_groups TO public"
    gConnection.Execute "GRANT all ON pgadmin_groups TO public"
    EndMsg
  End If
  If ObjectExists("pgadmin_indexes", tTable) <> 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Dropping corrupted pgAdmin Indexes View..."
    LogMsg "Executing: DROP TABLE pgadmin_indexes"
    gConnection.Execute "DROP TABLE pgadmin_indexes"
    EndMsg
  End If
  If ObjectExists("pgadmin_indexes", tView) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Indexes View..."
    LogMsg "Executing: " & SQL_PGADMIN_INDEXES
    gConnection.Execute SQL_PGADMIN_INDEXES
    LogMsg "Executing: GRANT all ON pgadmin_indexes TO public"
    gConnection.Execute "GRANT all ON pgadmin_indexes TO public"
    EndMsg
  End If
  If ObjectExists("pgadmin_languages", tTable) <> 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Dropping corrupted pgAdmin Languages View..."
    LogMsg "Executing: DROP TABLE pgadmin_languages"
    gConnection.Execute "DROP TABLE pgadmin_languages"
    EndMsg
  End If
  If ObjectExists("pgadmin_languages", tView) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Languages View..."
    LogMsg "Executing: " & SQL_PGADMIN_LANGUAGES
    gConnection.Execute SQL_PGADMIN_LANGUAGES
    LogMsg "Executing: GRANT all ON pgadmin_languages TO public"
    gConnection.Execute "GRANT all ON pgadmin_languages TO public"
    EndMsg
  End If
  If ObjectExists("pgadmin_sequences", tTable) <> 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Dropping corrupted pgAdmin Sequences View..."
    LogMsg "Executing: DROP TABLE pgadmin_sequences"
    gConnection.Execute "DROP TABLE pgadmin_sequences"
    EndMsg
  End If
  If ObjectExists("pgadmin_sequences", tView) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Sequences View..."
    LogMsg "Executing: " & SQL_PGADMIN_SEQUENCES
    gConnection.Execute SQL_PGADMIN_SEQUENCES
    LogMsg "Executing: GRANT all ON pgadmin_sequences TO public"
    gConnection.Execute "GRANT all ON pgadmin_sequences TO public"
    EndMsg
  End If
  If ObjectExists("pgadmin_tables", tTable) <> 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Dropping corrupted pgAdmin Tables View..."
    LogMsg "Executing: DROP TABLE pgadmin_tables"
    gConnection.Execute "DROP TABLE pgadmin_tables"
    EndMsg
  End If
  If ObjectExists("pgadmin_tables", tView) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Tables View..."
    LogMsg "Executing: " & SQL_PGADMIN_TABLES
    gConnection.Execute SQL_PGADMIN_TABLES
    LogMsg "Executing: GRANT all ON pgadmin_tables TO public"
    gConnection.Execute "GRANT all ON pgadmin_tables TO public"
    EndMsg
  End If
  If ObjectExists("pgadmin_users", tTable) <> 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Dropping corrupted pgAdmin Users View..."
    LogMsg "Executing: DROP TABLE pgadmin_users"
    gConnection.Execute "DROP TABLE pgadmin_users"
    EndMsg
  End If
  If ObjectExists("pgadmin_users", tView) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Users View..."
    LogMsg "Executing: " & SQL_PGADMIN_USERS
    gConnection.Execute SQL_PGADMIN_USERS
    LogMsg "Executing: GRANT all ON pgadmin_users TO public"
    gConnection.Execute "GRANT all ON pgadmin_users TO public"
    EndMsg
  End If
  If ObjectExists("pgadmin_triggers", tTable) <> 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Dropping corrupted pgAdmin Triggers View..."
    LogMsg "Executing: DROP TABLE pgadmin_triggers"
    gConnection.Execute "DROP TABLE pgadmin_triggers"
    EndMsg
  End If
  If ObjectExists("pgadmin_triggers", tView) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Triggers View..."
    LogMsg "Executing: " & SQL_PGADMIN_TRIGGERS
    gConnection.Execute SQL_PGADMIN_TRIGGERS
    LogMsg "Executing: GRANT all ON pgadmin_triggers TO public"
    gConnection.Execute "GRANT all ON pgadmin_triggers TO public"
    EndMsg
  End If
  If ObjectExists("pgadmin_views", tTable) <> 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Dropping corrupted pgAdmin Views View..."
    LogMsg "Executing: DROP TABLE pgadmin_views"
    gConnection.Execute "DROP TABLE pgadmin_views"
    EndMsg
  End If
  If ObjectExists("pgadmin_views", tView) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin Views View..."
    LogMsg "Executing: " & SQL_PGADMIN_VIEWS
    gConnection.Execute SQL_PGADMIN_VIEWS
    LogMsg "Executing: GRANT all ON pgadmin_views TO public"
    gConnection.Execute "GRANT all ON pgadmin_views TO public"
    EndMsg
  End If
  
  If ObjectExists("pgadmin_dev_functions", tTable) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin_dev_functions table..."
    LogMsg "Executing: " & SQL_PGADMIN_DEV_FUNCTIONS
    gConnection.Execute SQL_PGADMIN_DEV_FUNCTIONS
    LogMsg "Executing: GRANT all ON pgadmin_dev_functions TO public"
    gConnection.Execute "GRANT all ON pgadmin_dev_functions TO public"
    cmp_Project_Move_Functions "pgadmin_temp_functions", "", "pgadmin_dev_functions"
    cmp_Table_DropIfExists "pgadmin_temp_functions"
    EndMsg
  End If
  
  If ObjectExists("pgadmin_dev_triggers", tTable) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin_dev_triggers table..."
    LogMsg "Executing: " & SQL_PGADMIN_DEV_TRIGGERS
    gConnection.Execute SQL_PGADMIN_DEV_TRIGGERS
    LogMsg "Executing: GRANT all ON pgadmin_dev_triggers TO public"
    gConnection.Execute "GRANT all ON pgadmin_dev_triggers TO public"
    cmp_Project_Move_Triggers "pgadmin_temp_triggers", "", "pgadmin_dev_triggers"
    cmp_Table_DropIfExists "pgadmin_temp_triggers"
    EndMsg
  End If
  
  If ObjectExists("pgadmin_dev_views", tTable) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin_dev_triggers table..."
    LogMsg "Executing: " & SQL_PGADMIN_DEV_VIEWS
    gConnection.Execute SQL_PGADMIN_DEV_VIEWS
    LogMsg "Executing: GRANT all ON pgadmin_dev_views TO public"
    gConnection.Execute "GRANT all ON pgadmin_dev_views TO public"
    cmp_Project_Move_Views "pgadmin_temp_views", "", "pgadmin_dev_views"
    cmp_Table_DropIfExists "pgadmin_temp_views"
    EndMsg
  End If
  
  If ObjectExists("pgadmin_dev_dependencies", tTable) = 0 Then
    If Not SuperuserChk Then Exit Sub
    StartMsg "Creating pgAdmin_dev_triggers table..."
    LogMsg "Executing: " & SQL_PGADMIN_DEV_DEPENDENCIES
    gConnection.Execute SQL_PGADMIN_DEV_DEPENDENCIES
    LogMsg "Executing: GRANT all ON pgadmin_dev_dependencies TO public"
    gConnection.Execute "GRANT all ON pgadmin_dev_dependencies TO public"
    EndMsg
  End If
    
  'Set the SSO Version on the server
  LogMsg "Executing: UPDATE pgadmin_param SET param_value = '" & Val(SSO_VERSION) & "' WHERE param_id = 1"
  gConnection.Execute "UPDATE pgadmin_param SET param_value = '" & Val(SSO_VERSION) & "' WHERE param_id = 1"
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "basSQL, Chk_HelperObjects"
End Sub

Public Sub Drop_Objects(Optional bDropAll As Boolean)
On Error Resume Next
  StartMsg "Dropping pgAdmin Server Side Objects..."

  'Drop temp tables
  LogMsg "Executing: DROP TABLE pgadmin_temp_functions"
  gConnection.Execute "DROP TABLE pgadmin_temp_functions"
  
  LogMsg "Executing: DROP TABLE pgadmin_temp_triggers"
  gConnection.Execute "DROP TABLE pgadmin_temp_triggers"
  
  LogMsg "Executing: DROP TABLE pgadmin_temp_views"
  gConnection.Execute "DROP TABLE pgadmin_temp_views"
  
  'Backup dev objects to temp tables
  cmp_Project_Move_Functions "pgadmin_dev_functions", "", "pgadmin_temp_functions"
  cmp_Project_Move_Triggers "pgadmin_dev_triggers", "", "pgadmin_temp_triggers"
  cmp_Project_Move_Views "pgadmin_dev_views", "", "pgadmin_temp_views"

 
  'Drop tables
  If bDropAll = True Then
    LogMsg "Executing: DROP TABLE pgadmin_sys"
    gConnection.Execute "DROP TABLE pgadmin_sys"
    LogMsg "Executing: DROP TABLE pgadmin_log"
    gConnection.Execute "DROP TABLE pgadmin_log"
  End If
  LogMsg "Executing: DROP TABLE pgadmin_seq_cache"
  gConnection.Execute "DROP TABLE pgadmin_seq_cache"
  LogMsg "Executing: DROP TABLE pgadmin_table_cache"
  gConnection.Execute "DROP TABLE pgadmin_table_cache"
  
  'Drop functions
  LogMsg "Executing: DROP FUNCTION pgadmin_get_desc(oid)"
  gConnection.Execute "DROP FUNCTION pgadmin_get_desc(oid)"
  LogMsg "Executing: DROP FUNCTION pgadmin_get_col_def(oid, int4)"
  gConnection.Execute "DROP FUNCTION pgadmin_get_col_def(oid, int4)"
  LogMsg "Executing: DROP FUNCTION pgadmin_get_handler(oid)"
  gConnection.Execute "DROP FUNCTION pgadmin_get_handler(oid)"
  LogMsg "Executing: DROP FUNCTION pgadmin_get_type(oid)"
  gConnection.Execute "DROP FUNCTION pgadmin_get_type(oid)"
  LogMsg "Executing: DROP FUNCTION pgadmin_get_rows(oid)"
  gConnection.Execute "DROP FUNCTION pgadmin_get_rows(oid)"
  LogMsg "Executing: DROP FUNCTION pgadmin_get_sequence(oid)"
  gConnection.Execute "DROP FUNCTION pgadmin_get_sequence(oid)"
  LogMsg "Executing: DROP FUNCTION pgadmin_get_function_name(oid)"
  gConnection.Execute "DROP FUNCTION pgadmin_get_function_name(oid)"
  LogMsg "Executing: DROP FUNCTION pgadmin_get_function_arguments(oid)"
  gConnection.Execute "DROP FUNCTION pgadmin_get_function_arguments(oid)"
  LogMsg "Executing: DROP FUNCTION pgadmin_get_function_name(oid)"
  gConnection.Execute "DROP FUNCTION pgadmin_get_function_name(oid)"
  LogMsg "Executing: DROP FUNCTION pgadmin_get_function_arguments(oid)"
  gConnection.Execute "DROP FUNCTION pgadmin_get_function_arguments(oid)"
  
  'Drop views
  LogMsg "Executing: DROP VIEW pgadmin_checks"
  gConnection.Execute "DROP VIEW pgadmin_checks"
  LogMsg "Executing: DROP VIEW pgadmin_databases"
  gConnection.Execute "DROP VIEW pgadmin_databases"
  LogMsg "Executing: DROP VIEW pgadmin_functions"
  gConnection.Execute "DROP VIEW pgadmin_functions"
  LogMsg "Executing: DROP VIEW pgadmin_groups"
  gConnection.Execute "DROP VIEW pgadmin_groups"
  LogMsg "Executing: DROP VIEW pgadmin_indexes"
  gConnection.Execute "DROP VIEW pgadmin_indexes"
  LogMsg "Executing: DROP VIEW pgadmin_languages"
  gConnection.Execute "DROP VIEW pgadmin_languages"
  LogMsg "Executing: DROP VIEW pgadmin_sequences"
  gConnection.Execute "DROP VIEW pgadmin_sequences"
  LogMsg "Executing: DROP VIEW pgadmin_tables"
  gConnection.Execute "DROP VIEW pgadmin_tables"
  LogMsg "Executing: DROP VIEW pgadmin_triggers"
  gConnection.Execute "DROP VIEW pgadmin_triggers"
  LogMsg "Executing: DROP VIEW pgadmin_users"
  gConnection.Execute "DROP VIEW pgadmin_users"
  LogMsg "Executing: DROP VIEW pgadmin_views"
  gConnection.Execute "DROP VIEW pgadmin_views"
  
  'Drop all views as tables incase psql has re-written them as such
  LogMsg "Executing: DROP TABLE pgadmin_checks"
  gConnection.Execute "DROP TABLE pgadmin_checks"
  LogMsg "Executing: DROP TABLE pgadmin_databases"
  gConnection.Execute "DROP TABLE pgadmin_databases"
  LogMsg "Executing: DROP TABLE pgadmin_functions"
  gConnection.Execute "DROP TABLE pgadmin_functions"
  LogMsg "Executing: DROP TABLE pgadmin_indexes"
  gConnection.Execute "DROP TABLE pgadmin_indexes"
  LogMsg "Executing: DROP TABLE pgadmin_languages"
  gConnection.Execute "DROP TABLE pgadmin_languages"
  LogMsg "Executing: DROP TABLE pgadmin_sequences"
  gConnection.Execute "DROP TABLE pgadmin_sequences"
  LogMsg "Executing: DROP TABLE pgadmin_tables"
  gConnection.Execute "DROP TABLE pgadmin_tables"
  LogMsg "Executing: DROP TABLE pgadmin_triggers"
  gConnection.Execute "DROP TABLE pgadmin_triggers"
  LogMsg "Executing: DROP TABLE pgadmin_views"
  gConnection.Execute "DROP TABLE pgadmin_views"
  LogMsg "Executing: DROP TABLE pgadmin_dev_functions"
  gConnection.Execute "DROP TABLE pgadmin_dev_functions"
  LogMsg "Executing: DROP TABLE pgadmin_dev_triggers"
  gConnection.Execute "DROP TABLE pgadmin_dev_triggers"
  LogMsg "Executing: DROP TABLE pgadmin_dev_views"
  gConnection.Execute "DROP TABLE pgadmin_dev_views"
  LogMsg "Executing: DROP TABLE pgadmin_dev_dependencies"
  gConnection.Execute "DROP TABLE pgadmin_dev_dependencies"
  EndMsg
End Sub

Public Sub Update_TableCache()
On Error GoTo Err_Handler
Dim rsTables As New Recordset
Dim rsCount As New Recordset
  StartMsg "Updating Table Cache..."
  LogMsg "Executing: SELECT DISTINCT ON (table_oid) table_oid, table_name FROM pgadmin_tables"
  rsTables.Open "SELECT DISTINCT ON (table_oid) table_oid, table_name FROM pgadmin_tables", gConnection, adOpenForwardOnly
  LogMsg "Executing: DELETE FROM pgadmin_table_cache WHERE table_timestamp >= '" & Format(Date, "yyyy-mm-dd") & " 00:00:00' AND table_timestamp <= '" & Format(Date, "yyyy-mm-dd") & " 23:59:59'"
  gConnection.Execute "DELETE FROM pgadmin_table_cache WHERE table_timestamp >= '" & Format(Date, "yyyy-mm-dd") & " 00:00:00' AND table_timestamp <= '" & Format(Date, "yyyy-mm-dd") & " 23:59:59'"
  While Not rsTables.EOF
    If rsTables!table_name <> "pg_log" And rsTables!table_name <> "pg_variable" And rsTables!table_name <> "pg_xactlock" Then
      If rsCount.State <> adStateClosed Then rsCount.Close
      LogMsg "Executing: SELECT count(*) AS rows FROM " & QUOTE & rsTables!table_name & QUOTE
      rsCount.Open "SELECT count(*) AS rows FROM " & QUOTE & rsTables!table_name & QUOTE, gConnection, adOpenForwardOnly
      LogMsg "Executing: INSERT INTO pgadmin_table_cache(table_oid, table_rows) VALUES(" & rsTables!table_oid & ", " & rsCount!Rows & ")"
      gConnection.Execute "INSERT INTO pgadmin_table_cache(table_oid, table_rows) VALUES(" & rsTables!table_oid & ", " & rsCount!Rows & ")"
    End If
    rsTables.MoveNext
  Wend
  If rsTables.State <> adStateClosed Then rsTables.Close
  If rsCount.State <> adStateClosed Then rsCount.Close
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "basSQL, Update_TableCache"
End Sub

Public Sub Update_SequenceCache()
On Error GoTo Err_Handler
Dim rsSeqs As New Recordset
Dim rsData As New Recordset
Dim szIsCycled As String
  StartMsg "Updating Sequence Cache..."
  LogMsg "Executing: SELECT sequence_oid, sequence_name FROM pgadmin_sequences"
  rsSeqs.Open "SELECT sequence_oid, sequence_name FROM pgadmin_sequences", gConnection, adOpenForwardOnly
  LogMsg "Executing: DELETE FROM pgadmin_seq_cache WHERE sequence_timestamp >= '" & Format(Date, "yyyy-mm-dd") & " 00:00:00' AND sequence_timestamp <= '" & Format(Date, "yyyy-mm-dd") & " 23:59:59'"
  gConnection.Execute "DELETE FROM pgadmin_seq_cache WHERE sequence_timestamp >= '" & Format(Date, "yyyy-mm-dd") & " 00:00:00' AND sequence_timestamp <= '" & Format(Date, "yyyy-mm-dd") & " 23:59:59'"
  While Not rsSeqs.EOF
    If rsData.State <> adStateClosed Then rsData.Close
    LogMsg "Executing: SELECT * FROM " & QUOTE & rsSeqs!sequence_name & QUOTE
    rsData.Open "SELECT * FROM " & QUOTE & rsSeqs!sequence_name & QUOTE, gConnection, adOpenForwardOnly
    If rsData!is_cycled = 1 Or rsData!is_cycled = True Then
      szIsCycled = "Yes"
    Else
      szIsCycled = "No"
    End If
    LogMsg "Executing: INSERT INTO pgadmin_seq_cache(sequence_oid, sequence_last_value, sequence_increment_by, sequence_max_value, sequence_min_value, sequence_cache_value, sequence_is_cycled) VALUES(" & rsSeqs!sequence_OID & ", " & rsData!last_value & ", " & rsData!increment_by & ", " & rsData!max_value & ", " & rsData!min_value & ", " & rsData!cache_value & ", '" & szIsCycled & "')"
    gConnection.Execute "INSERT INTO pgadmin_seq_cache(sequence_oid, sequence_last_value, sequence_increment_by, sequence_max_value, sequence_min_value, sequence_cache_value, sequence_is_cycled) VALUES(" & rsSeqs!sequence_OID & ", " & rsData!last_value & ", " & rsData!increment_by & ", " & rsData!max_value & ", " & rsData!min_value & ", " & rsData!cache_value & ", '" & szIsCycled & "')"
    rsSeqs.MoveNext
  Wend
  If rsSeqs.State <> adStateClosed Then rsSeqs.Close
  If rsData.State <> adStateClosed Then rsData.Close
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "basSQL, Update_SequenceCache"
End Sub

Public Sub CreateMSysConf()
On Error GoTo Err_Handler
  StartMsg "Creating MSysConf Table..."
  LogQuery "CREATE TABLE msysconf(Config " & QUOTE & "int" & QUOTE & " NOT NULL, chValue varchar(255), nValue " & QUOTE & "int4" & QUOTE & ", Comments varchar(255))"
  gConnection.Execute "CREATE TABLE msysconf(Config " & QUOTE & "int" & QUOTE & " NOT NULL, chValue varchar(255), nValue " & QUOTE & "int4" & QUOTE & ", Comments varchar(255))"
  LogMsg "Executing: INSERT INTO msysconf VALUES ('101', '', '1', 'Allow local storage of passwords in attachments')"
  gConnection.Execute "INSERT INTO msysconf VALUES ('101', '', '1', 'Allow local storage of passwords in attachments')"
  LogMsg "Executing: INSERT INTO msysconf VALUES ('102', '', '10', 'Background population delay')"
  gConnection.Execute "INSERT INTO msysconf VALUES ('102', '', '10', 'Background population delay')"
  LogMsg "Executing: INSERT INTO msysconf VALUES ('103', '', '100', 'Background population size')"
  gConnection.Execute "INSERT INTO msysconf VALUES ('103', '', '100', 'Background population size')"
  EndMsg
  MsgBox "The MSysConf table has been created with default values.", vbExclamation
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "basSQL, CreateMSysConf"
End Sub

Public Function SuperuserChk() As Boolean
On Error GoTo Err_Handler
  If SuperUser = False Then
    SuperuserChk = False
    ActionCancelled = True
    MsgBox "pgAdmin could not update it's Server Side Objects because you are not a superuser." & vbCrLf & vbCrLf & "Please login as a superuser.", vbExclamation, "Error"
  Else
    SuperuserChk = True
  End If
  Exit Function
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "basSQL, SuperuserChk"
End Function

Public Function RsExecuteGetResult(ByVal szQuery As String) As Variant
    Dim rsTemp As New Recordset
    
    If rsTemp.State <> adStateClosed Then rsTemp.Close
    LogMsg "Executing: " & szQuery
    rsTemp.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
    
    If Not (rsTemp.EOF) Then
        RsExecuteGetResult = rsTemp(0).Value
        rsTemp.Close
    End If
End Function
