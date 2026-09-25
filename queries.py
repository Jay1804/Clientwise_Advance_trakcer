"""SQL queries used by the Advance Tracker Client report app.

CASE_REPORT_QUERY is the client-wise case/check report. Two bugs in the
original hand-written query were fixed here so it actually runs on MySQL:

- The `ec_client` join alias ("ec") collided with the later `ec_checks` join,
  which also used alias "ec" (MySQL: "Not unique table/alias: 'ec'"). The
  ec_client join was renamed to "eclient".
- A missing comma between `ec.client_external_id` and `emc.Company_name` in
  the SELECT list has been added.
- The second "Last_Inerim_Report_Severity" column (for the *final* report)
  was renamed to "Last_Final_Report_Severity" so the two columns don't
  collide when loaded into a DataFrame/Excel sheet.
- contact_value/MOBILE_NO are built from a correlated subquery on
  ec_candidate_contact_details rather than a plain LEFT JOIN: a candidate can
  have several contact rows (one per PRIORITY), and joining directly would
  have duplicated every case row per extra contact. The subquery picks the
  candidate's PRIORITY=1 (lowest priority number) active contact instead.
"""

CLIENT_LOOKUP_QUERY = """
SELECT eclient.client_external_id, emc.company_name
FROM ec_client eclient
JOIN ec_master_company emc ON emc.company_id = eclient.client_id
WHERE eclient.client_external_id IS NOT NULL
ORDER BY emc.company_name
"""

CHECK_NAME_LOOKUP_QUERY = """
SELECT DISTINCT check_name
FROM ec_checks
WHERE check_name IS NOT NULL
ORDER BY check_name
"""

CHECK_NAME_LOOKUP_BY_CLIENT_QUERY = """
SELECT DISTINCT ec.check_name
FROM ec_case_checks ecc
JOIN ec_checks ec ON ecc.check_id = ec.check_id
JOIN ec_case_master ecm ON ecc.case_id = ecm.case_id
JOIN ec_master_company emc ON emc.company_id = ecm.client_id
JOIN ec_client eclient ON eclient.client_id = emc.company_id
WHERE ec.check_name IS NOT NULL
AND eclient.client_external_id IN :client_ids
ORDER BY ec.check_name
"""

# Options for these lookups are derived with the same translation functions
# and exclusions CASE_REPORT_QUERY itself uses, so a selected option is
# guaranteed to match rows the report can actually return. Both dedupe the
# small set of raw codes first and translate only that handful of values -
# ec_case_master/ec_case_checks are large enough that translating every row
# before deduping (i.e. `SELECT DISTINCT fn_case_status(case_status) FROM
# ec_case_master`) was observed to take 30s+; deduping first keeps it to ~2-3s.
CASE_STATUS_LOOKUP_QUERY = """
SELECT checkpoint_live.fn_case_status(codes.case_status) AS case_status
FROM (SELECT DISTINCT case_status FROM ec_case_master WHERE case_status NOT IN (8,14)) codes
ORDER BY case_status
"""

# check_severity is a free-text column with ~140 messy legacy values overall
# and no usable index - an unscoped DISTINCT scan over it was observed to
# take 50s+ regardless of any date scoping, so the underlying table scan is
# the bottleneck, not the row count. Scoping to specific clients (like
# CHECK_NAME_LOOKUP_BY_CLIENT_QUERY) sidesteps that: a client's own checks
# are a small enough slice to come back in ~2-3s with a small, clean set of
# values. With no client selected there's no cheap way to list it, so the UI
# falls back to typing the value directly or an explicit "load anyway" ask.
CHECK_SEVERITY_LOOKUP_QUERY = """
SELECT DISTINCT check_severity
FROM ec_case_checks
WHERE check_severity IS NOT NULL AND check_severity <> ''
ORDER BY check_severity
"""

CHECK_SEVERITY_LOOKUP_BY_CLIENT_QUERY = """
SELECT DISTINCT ecc.check_severity
FROM ec_case_checks ecc
JOIN ec_case_master ecm ON ecc.case_id = ecm.case_id
JOIN ec_master_company emc ON emc.company_id = ecm.client_id
JOIN ec_client eclient ON eclient.client_id = emc.company_id
WHERE ecc.check_severity IS NOT NULL AND ecc.check_severity <> ''
AND eclient.client_external_id IN :client_ids
ORDER BY ecc.check_severity
"""

CHECK_STATUS_LOOKUP_QUERY = """
SELECT checkpoint_live.fn_check_status(codes.check_status) AS check_status
FROM (SELECT DISTINCT check_status FROM ec_case_checks WHERE check_status <> 9) codes
ORDER BY check_status
"""

FLEX_FIELD_QUERY = """
SELECT eccf.client_id,
       eclient.client_external_id,
       emc.company_name AS 'Client Name',
       eccf.client_field_id,
       eccf.field_id,
       eccf.field_name,
       IF(eccf.`status` = 0, 'Passive', 'Active') AS 'Status'
FROM ec_client_case_fields eccf
LEFT JOIN ec_client eclient ON eccf.client_id = eclient.client_id
INNER JOIN ec_master_company emc ON eccf.client_id = emc.company_id
WHERE eccf.field_id BETWEEN 1 AND 30
"""

CASE_REPORT_QUERY = """
SELECT ecc.Case_Check_id,case_ars_no,
checkpoint_live.fn_case_data_source(ecm.data_source) AS 'Case_Data_Source',
eclient.client_external_id,
emc.Company_name,
CONCAT(first_name,' ',IFNULL(middle_name,''),' ',IFNULL(last_name,'')) AS Candidate_name,
Process_name,  received_date AS case_received_date,
ecm.created_date AS case_created_date,
ecff.CASE_FLEX_FIELD1,ecff.CASE_FLEX_FIELD2,ecff.CASE_FLEX_FIELD3,ecff.CASE_FLEX_FIELD4,
ecff.CASE_FLEX_FIELD5,ecff.CASE_FLEX_FIELD6,ecff.CASE_FLEX_FIELD7,ecff.CASE_FLEX_FIELD8,
ecff.CASE_FLEX_FIELD9,ecff.CASE_FLEX_FIELD10,ecff.CASE_FLEX_FIELD11,ecff.CASE_FLEX_FIELD12,
ecff.CASE_FLEX_FIELD13,ecff.CASE_FLEX_FIELD14,ecff.CASE_FLEX_FIELD15,ecff.CASE_FLEX_FIELD16,
ecff.CASE_FLEX_FIELD17,ecff.CASE_FLEX_FIELD18,ecff.CASE_FLEX_FIELD19,ecff.CASE_FLEX_FIELD20,
ecff.CASE_FLEX_FIELD21,ecff.CASE_FLEX_FIELD22,ecff.CASE_FLEX_FIELD23,ecff.CASE_FLEX_FIELD24,
ecff.CASE_FLEX_FIELD25,ecff.CASE_FLEX_FIELD26,ecff.CASE_FLEX_FIELD27,ecff.CASE_FLEX_FIELD28,
ecff.CASE_FLEX_FIELD29,ecff.CASE_FLEX_FIELD30,
checkpoint_live.fn_case_status(case_status) AS 'Case_status',
checkpoint_live.fn_check_status(check_status) AS 'check_status',
(CASE checkpoint_live.fn_check_status(check_status)
WHEN 'Completed' THEN 'Completed'
WHEN 'case closed by client' THEN 'case closed by client'
WHEN 'Closed with Insufficiency' THEN 'Closed with Insufficiency'
WHEN 'Awaiting Response' THEN 'Work in Progress'
WHEN 'Work in Progress' THEN 'Work in Progress'
WHEN 'On Hold' THEN 'On Hold'
WHEN 'Insufficient' THEN 'Insufficient'
WHEN 'Escalated' THEN 'Work in Progress'
WHEN 'New/UnAssigned' THEN 'Work in Progress'
WHEN 'In Research' THEN 'Work in Progress'
ELSE checkpoint_live.fn_check_status(check_status) END) AS 'check_status1',
ecc.check_name AS 'check_name',
ec.check_name AS 'Check_unique_name',
(CASE ecc.VER_TYPE
WHEN 1 THEN 'Verbal'
WHEN 2 THEN 'Written'
WHEN 3 THEN 'VRWP' END) AS 'Ver_Type',
check_disposition_id,disposition_name,check_severity,
REPLACE(ecc.closure_comments,'rn','') AS closure_comments,ecc.check_closure_date,
ecc.insuff_remarks,

@insuffdate:=(SELECT action_taken_on FROM ec_case_history ech WHERE ecc.case_check_id=ech.check_id
AND action_taken in (' case Status changed to : Case Insufficient',
' case Status changed to : InSufficient',
'Insuff Raised',
'Marked Insufficient',
'Marked Insufficient (Parallel Research)',
'New Status - Case Insufficient',
'New Status - InSufficient',
'Vendor request closed & Insuff raised',
'Case Insufficient',
'Check Created | Marked Insufficient',
'Check Updated | Marked Insufficient',
'New Status - Case Insuff Updated',
'New Status - New Case | Case Insufficient',
'Check Insuff raised',
'Case level Insuff raised',
'New Status - Case Insufficiency raised',
'Insufficient - Rework on Report',
'Insuff accepted' ) ORDER BY action_id LIMIT 1) AS 'First_Insuff_Date',

ecc.insuff_fulfill_date,
office_name location,
IF(ecc.family_id=4,emei.institute_name,IF(ecc.family_id=5,emc1.company_name,emcy.city_name)) AS verification_source,
ecm.case_expected_closure_date case_due_date,check_created_on,ecc.go_ahead_date,ecc.copy_of_check,

IFNULL(ec.CHECK_OPS_NAME,LEFT(REPLACE(family_name,' Family',''),3)) AS 'Check Ops Name',

ecc.reopen_date AS 'check_reopen_date',ecm.reopen_date AS 'case_reopen_date',
checkpoint_live.fn_ver_summary(ecc.VER_SUMMARY) AS Ver_Summary,
checkpoint_live.fn_ver_procedure(ecc.VERIFICATION_PROCEDURE) AS VERIFICATION_PROCEDURE,
checkpoint_live.fn_check_sub_status(ecc.sub_status) AS 'Check Sub Status',
(SELECT action_taken_on FROM ec_case_history ech
WHERE ecc.case_check_id=ech.check_id AND action_taken='Insuff Qc Status Updateas'
AND action_comments='Accepted' ORDER BY action_id DESC LIMIT 1) AS 'Inusff Accepted',

(SELECT report_sent_on FROM ec_case_reports ecr WHERE ecr.case_id=ecm.case_id AND report_type=0 AND report_status=5 ORDER BY case_report_id DESC LIMIT 1) 'Last_Inerim_Report_Sent_Date',
(SELECT report_severity FROM ec_case_reports ecr WHERE ecr.case_id=ecm.case_id AND report_type=0 AND report_status=5 ORDER BY case_report_id DESC LIMIT 1) 'Last_Inerim_Report_Severity',
(SELECT report_sent_on FROM ec_case_reports ecr WHERE ecr.case_id=ecm.case_id AND report_type=0 AND report_status=5 ORDER BY case_report_id ASC LIMIT 1) 'First_interim_sent_Date',
(SELECT report_severity FROM ec_case_reports ecr WHERE ecr.case_id=ecm.case_id AND report_type=0 AND report_status=5 ORDER BY case_report_id ASC LIMIT 1) 'First_interim_sent_Date_Sevirity',
(SELECT report_sent_on FROM ec_case_reports ecr WHERE ecr.case_id=ecm.case_id AND report_type=1 AND report_status=5 ORDER BY case_report_id DESC LIMIT 1) 'Last_Final_Report_Sent_Date',
(SELECT report_severity FROM ec_case_reports ecr WHERE ecr.case_id=ecm.case_id AND report_type=1 AND report_status=5 ORDER BY case_report_id DESC LIMIT 1) 'Last_Final_Report_Severity',
(SELECT report_sent_on FROM ec_case_reports ecr WHERE ecr.case_id=ecm.case_id AND report_type=2 AND report_status=5 ORDER BY case_report_id DESC LIMIT 1) 'Last_Additional_Report_Sent_Date',
(SELECT report_severity FROM ec_case_reports ecr WHERE ecr.case_id=ecm.case_id AND report_type=2 AND report_status=5 ORDER BY case_report_id DESC LIMIT 1) 'Last_Additional_Report_Severity',
dqc_released_date,
(CASE TIER
WHEN 0 THEN 'Overseas'
WHEN 1 THEN 'Tier 1'
WHEN 2 THEN 'Tier 2'
WHEN 3 THEN 'Tier 3'
WHEN 4 THEN 'Tier 4' END ) AS 'Tier',

CONCAT(eud1.user_first_name,' ',eud1.user_last_name) AS 'DS Name',
QUEUE_NAME,
IF(@insuffdate is not null,if(@insuffdate<=dqc_released_date,'L1','L2'),'') AS 'Insuff Type',
IF(@insuffdate is not null,if(check_status in (0,1,2,3,4,5,6,7,13),'WIP','Non-Wip'),'Others') AS 'Insuff WIP Type',
PRIORITIZED_REQUESTED_EDC as 'EDC Prioritized Requested Date',
PRIORITIZED_REVISED_EDC as 'EDC Prioritized Revised Date',
if(ecc.check_status in (8,10,11,12),fn_user_name(VQC_REVIEWER_ID),'') as 'VQC done by',

@contact_value:=(SELECT eccd.CONTACT_VALUE FROM ec_candidate_contact_details eccd
WHERE eccd.candidate_id=ecm.candidate_id AND eccd.status=1
ORDER BY eccd.priority ASC LIMIT 1) AS contact_value,
IF(LENGTH(ecc1.MOBILE_NO)=0,@contact_value,CONCAT('con ',ecc1.MOBILE_NO)) AS MOBILE_NO,
ecc1.EMAIL_ADDRESS AS EMAIL_ADDRESS

FROM ec_case_master ecm
LEFT JOIN ec_case_fields ecff ON ecm.case_id=ecff.case_id

LEFT JOIN ec_case_checks ecc ON ecm.case_id = ecc.case_id
LEFT JOIN ec_check_queues ecq ON ecc.check_queue=ecq.queue_id AND ecc.check_id=ecq.check_id
LEFT JOIN ec_master_company emc ON emc.company_id = ecm.client_id
LEFT JOIN ec_client eclient ON eclient.client_id= emc.company_id
LEFT JOIN ec_client_process ecp ON ecm.process_id=ecp.process_id
LEFT JOIN ec_case_candidates ecc1 ON ecc1.candidate_id=ecm.candidate_id
LEFT JOIN ec_master_company_locations emcl ON ecm.client_office_id=emcl.office_id
LEFT JOIN ec_user_details eud ON ecc.check_verifier=eud.user_id
LEFT JOIN ec_user_details eud1 ON ecm.documented_by=eud1.user_id
LEFT JOIN ec_case_check_verification_source eccvs ON ecc.case_check_id=eccvs.case_check_id
LEFT JOIN ec_master_company emc1 ON eccvs.org_id=emc1.company_id
LEFT JOIN ec_master_educational_institute emei ON eccvs.org_id=emei.institute_id
LEFT JOIN ec_master_city emcy ON emcy.city_id=eccvs.org_id
LEFT JOIN ec_master_state ems ON emcy.state_id=ems.state_id
LEFT JOIN ec_checks ec ON ecc.check_id=ec.check_id
left join ec_check_families ecf on ec.family_id=ecf.family_id
LEFT JOIN ec_master_disposition emd ON ecc.check_disposition_id=emd.disposition_id
WHERE case_status NOT IN (8,14)
AND check_status <> 9
and ecc.check_id<>193
AND received_date BETWEEN :from_date AND :to_date
{client_filter_clause}
{check_name_filter_clause}
{case_status_filter_clause}
{check_status_filter_clause}
{check_severity_filter_clause}
"""

# --- Case-wise vs check-wise column classification ------------------------
#
# CASE_REPORT_QUERY is one row per (case, check) pair - a case with 3 checks
# produces 3 rows, all sharing the same case-level values. Every column in its
# SELECT list falls into exactly one of these two buckets, resolved against
# the actual DB schema (SHOW COLUMNS on each joined table) rather than guessed
# from naming alone - e.g. 'location' looks check-ish but is office_name from
# ec_master_company_locations, joined via ecm.client_office_id, so it's
# case-level; 'Tier' looks case-ish but is TIER from ec_master_city, joined
# via the check's own verification-source city, so it's check-level.
#
# - CASE_LEVEL_COLUMNS: from ec_case_master (keyed by CASE_ARS_NO) or another
#   table joined by case_id/candidate_id - repeats across every check row of
#   that case.
# - CHECK_LEVEL_COLUMNS: from ec_case_checks (keyed by CASE_CHECK_ID, exposed
#   as Case_Check_id) or another table joined by case_check_id/check_id -
#   varies row-by-row even within the same case.
#
# This is what the "Customize Columns" pivot-style aggregation in app.py
# relies on: hiding every check-level column collapses the report from
# check-level to case-level, since drop_duplicates() then only sees the
# case-level columns that are left.
CASE_LEVEL_COLUMNS = frozenset({
    "case_ars_no", "Case_Data_Source", "client_external_id", "Company_name",
    "Candidate_name", "Process_name", "case_received_date", "case_created_date",
    *(f"CASE_FLEX_FIELD{i}" for i in range(1, 31)),
    "Case_status", "location", "case_due_date", "case_reopen_date",
    "Last_Inerim_Report_Sent_Date", "Last_Inerim_Report_Severity",
    "First_interim_sent_Date", "First_interim_sent_Date_Sevirity",
    "Last_Final_Report_Sent_Date", "Last_Final_Report_Severity",
    "Last_Additional_Report_Sent_Date", "Last_Additional_Report_Severity",
    "dqc_released_date", "DS Name", "EDC Prioritized Requested Date",
    "EDC Prioritized Revised Date", "contact_value", "MOBILE_NO", "EMAIL_ADDRESS",
})

CHECK_LEVEL_COLUMNS = frozenset({
    "Case_Check_id", "check_status", "check_status1", "check_name",
    "Check_unique_name", "Ver_Type", "check_disposition_id", "disposition_name",
    "check_severity", "closure_comments", "check_closure_date", "insuff_remarks",
    "First_Insuff_Date", "insuff_fulfill_date", "verification_source",
    "check_created_on", "go_ahead_date", "copy_of_check", "Check Ops Name",
    "check_reopen_date", "Ver_Summary", "VERIFICATION_PROCEDURE",
    "Check Sub Status", "Inusff Accepted", "Tier", "QUEUE_NAME", "Insuff Type",
    "Insuff WIP Type", "VQC done by",
})

CLIENT_FILTER_CLAUSE = "AND eclient.client_external_id IN :client_ids"
CHECK_NAME_FILTER_CLAUSE = "AND ec.check_name IN :check_names"
# These filter on the same translated expressions CASE_REPORT_QUERY's SELECT
# list uses for Case_status/check_status, so filter values line up with what
# the report actually displays for those columns rather than raw DB codes.
CASE_STATUS_FILTER_CLAUSE = "AND checkpoint_live.fn_case_status(case_status) IN :case_statuses"
CHECK_STATUS_FILTER_CLAUSE = "AND checkpoint_live.fn_check_status(check_status) IN :check_statuses"
CHECK_SEVERITY_FILTER_CLAUSE = "AND check_severity IN :check_severities"

# --- Antecedent fields (ec_check_fields / ec_case_check_data) -----------------
#
# A case_check_id can have many antecedent field rows (Employee ID, Designation,
# Salary, etc. for an employment check; different fields entirely for other
# check families), so joining ec_case_check_data straight into CASE_REPORT_QUERY
# would multiply every check's row once per field it has data for. Instead this
# is kept as a separate, opt-in lookup: pick specific field name(s) via
# FIELD_NAME_LOOKUP_QUERY, then pull just those via ANTECEDENT_DATA_QUERY, both
# scoped to the Case_Check_id values already present in the fetched report, and
# the app pivots/merges the result in as extra columns.
#
# FIELD_NAME_LOOKUP_QUERY was originally scoped by client (joining through
# ec_case_checks/ec_case_master/ec_master_company/ec_client), which was fast
# for one client (~2-3s) but timed out at 60s+ once a report spanned more than
# a handful of clients - that join doesn't scale with client count. Scoping
# directly by the report's own Case_Check_id values instead (same set used by
# ANTECEDENT_DATA_QUERY below) skips that join chain entirely and was measured
# at ~3-6s even for 5,000-20,000 ids, regardless of how many clients they span.
#
# The same FIELD_NAME can appear under several FIELD_IDs (one per check
# version/family it was configured on), so lookups and the data pull both match
# on the field's name rather than its id.
FIELD_NAME_LOOKUP_QUERY = """
SELECT DISTINCT ecf.FIELD_NAME AS field_name
FROM ec_case_check_data eccd
JOIN ec_check_fields ecf ON eccd.FIELD_ID = ecf.FIELD_ID
WHERE eccd.CASE_CHECK_ID IN :case_check_ids
AND ecf.FIELD_NAME IS NOT NULL AND ecf.FIELD_NAME <> ''
ORDER BY ecf.FIELD_NAME
"""

ANTECEDENT_DATA_QUERY = """
SELECT eccd.CASE_CHECK_ID AS case_check_id,
       ecf.FIELD_NAME AS field_name,
       eccd.STATED_DATA AS stated_data,
       eccd.VERIFIED_DATA AS verified_data,
       eccd.DATA_UPDATED_ON AS data_updated_on
FROM ec_case_check_data eccd
JOIN ec_check_fields ecf ON eccd.FIELD_ID = ecf.FIELD_ID
WHERE ecf.FIELD_NAME IN :field_names
AND eccd.CASE_CHECK_ID IN :case_check_ids
"""
