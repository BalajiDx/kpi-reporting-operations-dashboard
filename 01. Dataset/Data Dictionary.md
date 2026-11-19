# Data Dictionary

Synthetic data, seed 61. Jul 2023 - Dec 2024. Client names are invented.

## Jobs
| Column | Description |
|---|---|
| JobID | Unique job |
| Client | Publisher client |
| Team | Data Conversion, Academic Editing, Typesetting |
| JobType | Each job type belongs to one team |
| ReceivedDate, DueDate, DeliveredDate | DeliveredDate blank if not delivered by 31 Dec 2024 |
| Pages | Pages in the job |
| QAErrors | Errors found by quality assurance |
| Rework | 1 if the job was sent back for rework |
| Status | Delivered, In progress, Needs review (bad dates found in cleaning) |

## Staffing
| Column | Description |
|---|---|
| Month | yyyy-mm |
| Team |  |
| FTE | Actual full-time equivalents |
| PlannedFTE | Planned |

## Escalations
| Column | Description |
|---|---|
| EscalationID |  |
| Date |  |
| Client, Team, JobID |  |
| Reason | Late delivery or Quality issue |
