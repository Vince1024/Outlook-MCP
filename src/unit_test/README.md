# MCP-OUTLOOK Unit Tests

## Installation

Install test dependencies:

```bash
pip install -r src/unit_test/requirements.txt
```

## Running Tests

### All tests
```bash
pytest src/unit_test/ -v
```

### Specific test file
```bash
pytest src/unit_test/email/test_send_email.py -v
```

### With code coverage
```bash
pytest src/unit_test/ -v --cov=src --cov-report=html
```

### Single test
```bash
pytest src/unit_test/email/test_send_email.py::TestSendEmail::test_send_email_basic -v
```

## Test Structure

```
src/unit_test/
├── __init__.py
├── conftest.py                          # Pytest configuration
├── requirements.txt                     # Test dependencies
├── README.md                            # This file
├── email/                               # 5 email tool tests
│   ├── __init__.py
│   ├── test_send_email.py
│   ├── test_get_inbox_emails.py
│   ├── test_get_sent_emails.py
│   ├── test_search_emails.py
│   └── test_create_draft_email.py
├── folder/                              # 3 folder tool tests
│   ├── __init__.py
│   ├── test_list_outlook_folders.py
│   ├── test_search_emails_in_custom_folder.py
│   └── test_list_outlook_rules.py
├── calendar/                            # 3 calendar tool tests
│   ├── __init__.py
│   ├── test_get_calendar_events.py
│   ├── test_create_calendar_event.py
│   └── test_search_calendar_events.py
└── contact/                             # 3 contact tool tests
    ├── __init__.py
    ├── test_get_contacts.py
    ├── test_create_contact.py
    └── test_search_contacts.py
```

**Total: 14 test files covering all 14 tools (100% coverage)**

## Available Tests

### Email Tools (5 tests)

#### test_send_email.py
- ✅ `test_send_email_basic`: Basic email sending
- ✅ `test_send_email_with_cc_bcc`: With CC and BCC recipients
- ✅ `test_send_email_with_importance`: With high/low importance
- ✅ `test_send_email_with_html_body`: HTML body content
- ✅ `test_send_email_with_signature`: With Outlook signature
- ✅ `test_send_email_failure`: Error handling
- ✅ `test_send_email_multiple_recipients`: Multiple recipients

#### test_get_inbox_emails.py
- ✅ `test_get_inbox_emails_basic`: Basic inbox retrieval
- ✅ `test_get_inbox_emails_unread_only`: Filter unread emails
- ✅ `test_get_inbox_emails_empty_inbox`: Handle empty inbox
- ✅ `test_get_inbox_emails_limit_enforced`: Limit enforcement
- ✅ `test_get_inbox_emails_failure`: Error handling
- ✅ `test_get_inbox_emails_with_format_error`: Handle format errors

#### test_get_sent_emails.py
- ✅ `test_get_sent_emails_basic`: Basic sent emails retrieval
- ✅ `test_get_sent_emails_empty_sent_folder`: Handle empty folder
- ✅ `test_get_sent_emails_limit_enforced`: Limit enforcement
- ✅ `test_get_sent_emails_failure`: Error handling
- ✅ `test_get_sent_emails_with_format_error`: Handle format errors
- ✅ `test_get_sent_emails_custom_limit`: Custom limit parameter

#### test_search_emails.py
- ✅ `test_search_emails_inbox_basic`: Search in inbox
- ✅ `test_search_emails_in_sent_folder`: Search in sent folder
- ✅ `test_search_emails_all_folders`: Search across all folders
- ✅ `test_search_emails_no_results`: Handle no results
- ✅ `test_search_emails_limit_enforced`: Limit enforcement
- ✅ `test_search_emails_failure`: Error handling
- ✅ `test_search_emails_with_format_error`: Handle format errors

#### test_create_draft_email.py
- ✅ `test_create_draft_email_basic`: Basic draft creation
- ✅ `test_create_draft_email_with_cc_bcc`: With CC and BCC
- ✅ `test_create_draft_email_with_html_body`: HTML body
- ✅ `test_create_draft_email_with_signature`: With signature
- ✅ `test_create_draft_email_with_signature_display_success`: Signature via Display()
- ✅ `test_create_draft_email_failure`: Error handling
- ✅ `test_create_draft_email_multiple_recipients`: Multiple recipients
- ✅ `test_create_draft_email_html_body_with_signature`: HTML + signature

### Folder Tools (3 tests)

#### test_list_outlook_folders.py
- ✅ `test_list_outlook_folders_basic`: Basic folder listing
- ✅ `test_list_outlook_folders_empty`: Handle empty folders
- ✅ `test_list_outlook_folders_with_excluded_stores`: Skip excluded stores
- ✅ `test_list_outlook_folders_with_store_error`: Handle store errors
- ✅ `test_list_outlook_folders_failure`: Error handling
- ✅ `test_list_outlook_folders_with_nested_folders`: Nested structure

#### test_search_emails_in_custom_folder.py
- ✅ `test_search_emails_in_custom_folder_basic`: Basic search
- ✅ `test_search_emails_folder_not_found`: Handle missing folder
- ✅ `test_search_emails_with_query`: Search with keyword
- ✅ `test_search_emails_with_days_back`: Date filter
- ✅ `test_search_emails_empty_folder`: Handle empty folder
- ✅ `test_search_emails_failure`: Error handling
- ✅ `test_search_emails_limit_enforced`: Limit enforcement

#### test_list_outlook_rules.py
- ✅ `test_list_outlook_rules_basic`: Basic rule listing
- ✅ `test_list_outlook_rules_empty`: Handle no rules
- ✅ `test_list_outlook_rules_with_from_condition`: From condition
- ✅ `test_list_outlook_rules_with_body_condition`: Body condition
- ✅ `test_list_outlook_rules_with_parse_error`: Handle parse errors
- ✅ `test_list_outlook_rules_failure`: Error handling
- ✅ `test_list_outlook_rules_with_nested_folder_path`: Nested folder paths

### Calendar Tools (3 tests)

#### test_get_calendar_events.py
- ✅ `test_get_calendar_events_basic`: Basic event retrieval
- ✅ `test_get_calendar_events_no_events`: Handle empty calendar
- ✅ `test_get_calendar_events_with_past`: Include past events
- ✅ `test_get_calendar_events_custom_days_ahead`: Custom days range
- ✅ `test_get_calendar_events_failure`: Error handling
- ✅ `test_get_calendar_events_with_format_error`: Handle format errors

#### test_create_calendar_event.py
- ✅ `test_create_calendar_event_basic`: Basic event creation
- ✅ `test_create_calendar_event_with_location`: With location
- ✅ `test_create_calendar_event_with_body`: With body content
- ✅ `test_create_calendar_event_with_attendees`: With attendees (meeting)
- ✅ `test_create_calendar_event_all_day`: All-day event
- ✅ `test_create_calendar_event_custom_reminder`: Custom reminder
- ✅ `test_create_calendar_event_invalid_date`: Invalid date handling
- ✅ `test_create_calendar_event_failure`: Error handling

#### test_search_calendar_events.py
- ✅ `test_search_calendar_events_by_subject`: Search by subject
- ✅ `test_search_calendar_events_by_location`: Search by location
- ✅ `test_search_calendar_events_no_results`: Handle no results
- ✅ `test_search_calendar_events_custom_days_range`: Custom range
- ✅ `test_search_calendar_events_case_insensitive`: Case-insensitive
- ✅ `test_search_calendar_events_failure`: Error handling
- ✅ `test_search_calendar_events_with_none_fields`: Handle None fields

### Contact Tools (3 tests)

#### test_get_contacts.py
- ✅ `test_get_contacts_basic`: Basic contact retrieval
- ✅ `test_get_contacts_with_search_filter`: Search filter
- ✅ `test_get_contacts_empty_folder`: Handle empty folder
- ✅ `test_get_contacts_limit_enforced`: Limit enforcement
- ✅ `test_get_contacts_failure`: Error handling
- ✅ `test_get_contacts_with_format_error`: Handle format errors

#### test_create_contact.py
- ✅ `test_create_contact_basic`: Basic contact creation
- ✅ `test_create_contact_with_company`: With company info
- ✅ `test_create_contact_with_phones`: With phone numbers
- ✅ `test_create_contact_complete`: All fields
- ✅ `test_create_contact_minimal`: Only required fields
- ✅ `test_create_contact_failure`: Error handling
- ✅ `test_create_contact_with_special_characters`: Special characters

#### test_search_contacts.py
- ✅ `test_search_contacts_by_name`: Search by name
- ✅ `test_search_contacts_by_email`: Search by email
- ✅ `test_search_contacts_by_company`: Search by company
- ✅ `test_search_contacts_no_results`: Handle no results
- ✅ `test_search_contacts_case_insensitive`: Case-insensitive
- ✅ `test_search_contacts_failure`: Error handling
- ✅ `test_search_contacts_with_none_fields`: Handle None fields
- ✅ `test_search_contacts_with_format_error`: Handle format errors

## Notes

- All tests use mocks to avoid accessing real Outlook
- Integration tests (that access real Outlook) are skipped by default
- To run integration tests: `pytest -k "integration" --runxfail`
- Tests cover all 14 MCP tools with comprehensive scenarios
- Tests are GDPR-compliant (no PII in logs)
- Each test file follows the same pattern for consistency

