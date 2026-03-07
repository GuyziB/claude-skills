---
name: weekly-task-allocation
description: Processes a Microsoft Planner Excel or CSV export to generate a structured, prioritised HTML email for weekly team task allocation. Use this skill whenever the user asks to generate a weekly team email, create a task allocation summary, or process tasks for the BMS or Controls team. Trigger this skill even if the user simply says "do the weekly email" or "process the planner file" — it is very likely they mean this workflow.
---

# Weekly Task Allocation Assistant

Turns a Microsoft Planner export into a formatted HTML team email, ready to paste into Outlook.

---

## Step 1: Data Input & Cleanup

Read the file at `C:\Users\guyboranyay\Downloads\Weekly Plan.xlsx`. If the file is not found, tell the user to place the Planner export there and stop.

Retain only these columns:

- `Bucket Name`
- `Priority`
- `Due Date`
- `Checklist Items`
- `Labels`
- `Description`
- `Task Name`

Discard all other columns.

---

## Step 2: Row Filtering

For every row, check `Labels` and `Priority`:

- **KEEP** if Priority is `Urgent`
- **KEEP** if Labels contains `ThisWeek`
- **DELETE** everything else

---

## Step 3: Extract General Team Notification

- Find the task named **"General Team Notification for Week Ahead"**
- Extract its `Description` — this becomes the email's opening body text
- Remove this row from the task tables (it is not a team task)
- If this task is not found, skip the notification block and note its absence

---

## Step 4: Organise by Team Member

Create one table per person, in this exact order:

| Order | Name    | Label to Match |
|-------|---------|----------------|
| 1     | Raymond | `Raymond`      |
| 2     | Jasmine | `JVT`          |
| 3     | Victor  | `Victor`       |
| 4     | Edward  | `Edward`       |
| 5     | Mark    | `Mark C`       |
| 6     | Guy     | `Guy`          |

A task belongs to a person if their label appears in that row's `Labels` column.

If a task has the label `Whole Team`, include it in **all** six tables.

---

## Step 5: Prioritise Tasks Within Each Table

Sort each person's tasks in this order:

1. **Highest:** `ThisWeek` label AND `Urgent` priority
2. **Middle:** `ThisWeek` label (any other priority)
3. **Standard:** `Important` priority (no `ThisWeek` label)

---

## Step 6: Generate HTML Email Output

Produce a complete, self-contained `.html` file with this structure:

**Subject line:** `Team Focus for Week [Current Week Number] - 2026`

**Body:**
1. "Good Morning Team, I hope you had a good weekend."
2. General Notification text (from Step 3)
3. One HTML table per team member showing `Bucket Name` and `Task Name`
4. Professional sign-off from Guy Baranyay

**Priority styling:**
- 🔴 `Urgent` + `ThisWeek` → bold red text or red row highlight
- 🟡 `ThisWeek` only → light yellow row highlight
- Standard → no special formatting

Tables should be clean, professional, and suitable for pasting into Outlook.

---

## Permanent Settings

| Setting       | Values |
|---------------|--------|
| Year          | 2026 |
| Team Labels   | Victor; Raymond; Mark C; JVT; Guy; Edward |
| Status Labels | Awaiting Client Response; Backlog; Closing; In Progress; Light Grey; Paused; Quotation; Update Required from Assignee; ThisWeek; Urgent; Whole Team |
