# Follow Up Machine

A cold-email follow-up autoresponder built on Google Apps Script, Gmail and a
spreadsheet. It replies into the original thread, stops the moment a lead
answers, and picks people back up when they go quiet again.

Everything lives in one spreadsheet. There is no server, no database and no
subscription.

---

## Contents

- [Your routine](#your-routine)
- [The two tracks](#the-two-tracks)
- [What runs on its own](#what-runs-on-its-own)
- [First-time setup](#first-time-setup)
- [Writing sequences](#writing-sequences)
- [Settings](#settings)
- [Statuses](#statuses)
- [What to do when](#what-to-do-when)
- [Limits](#limits)
- [Troubleshooting](#troubleshooting)
- [The files](#the-files)

---

## Your routine

Adding a lead takes about 30 seconds. Everything after this is automatic.

1. Type the address into **InboxScanner** column A, on a new row, leaving
   column B empty.
2. Run **Follow-Up System → Scan inbox for email**.
3. Tick **Enroll** on the correct thread — this is the thread every follow-up
   replies into, so pick the real cold email.
4. Fill in `leadName` on the new row in ActiveFollowUps. Without it
   `{{firstName}}` comes out blank.

That is the entire job. You never touch the lead again unless you want to.

> Enrol while the sending window is still open and the first follow-up goes out
> the same day. Otherwise it waits for tomorrow.

---

## The two tracks

### Track 1 — they never reply

```
enroll → step 1 → step 2 → step 3 → … → last step → Done
```

Fully automatic, no touches. The number of follow-ups is simply the number of
rows in that sequence's column. Want eight? Write eight rows.

### Track 2 — they reply, then go quiet

```
step 1 → step 2 → THEY REPLY → Replied  (sequence stops)
                                  ↓ silent for replyResumeDays
                       reEngage step 1 → 2 → 3 → … → Done
```

Also automatic, as long as `resumeOnReply` is ticked on that lead — which
happens by itself when `enrollDefaultResumeOnReply` is `TRUE`.

Reply again partway through the re-engage sequence and it pauses again, then
carries on **where it stopped** rather than restarting.

### These stack

A lead who replies once can receive the whole cold sequence *plus* the whole
re-engage sequence. Eight plus eight is sixteen emails to one person. Pick the
total deliberately — six to eight cold and four to five re-engage is a
reasonable ceiling. If twelve touches over two months has not moved someone,
the thirteenth is not the problem.

---

## What runs on its own

| Automatic | You |
|---|---|
| Sending every step on schedule | Enrolling a lead |
| Stopping the moment a lead replies | Writing the sequences |
| Resuming a lead who replied then went quiet | Actually replying to warm leads |
| Moving resumed leads onto the re-engage sequence | |
| Detecting bounces and out-of-office | |
| Respecting the daily cap and sending window | |
| Never emailing the same lead twice in one day | |
| Emailing you when something breaks | |

---

## First-time setup

1. **Enable the Gmail advanced service.** Apps Script → *Services* → add
   **Gmail API**. Without it every send fails. The system builds raw RFC
   messages so replies thread properly and always go to the lead rather than
   back to you.

2. **Run *Setup spreadsheet*** from the Follow-Up System menu. This creates all
   five sheets with headers, formats, dropdowns and checkboxes.

3. **Fill in Settings.** At minimum `timezone`, `notificationEmail` and the
   sending window. See [Settings](#settings).

4. **Write your sequences** in the Sequences sheet. See
   [Writing sequences](#writing-sequences).

5. **Set `enrollDefaultSequence`** to your main sequence name. Leave it blank
   and every enrolled lead starts with no sequence.

6. **Run *Create daily trigger*.** This builds three triggers: the daily run,
   the weekly summary on your `notificationDay`, and the checkbox watcher.
   Re-run it whenever you change `notificationDay` — the weekday is baked into
   the trigger when it is created.

7. **Send yourself a test.** *Send test email* renders exactly what a lead
   would receive, variables filled and image inline, with a `[TEST]` subject.
   Nothing is written to the sheet.

### The web app (optional)

`webapp/WebApp.gs` and `webapp/Index.html` give you a phone-friendly interface:
leads, a pipeline board, sequence editing, the send log and settings.

- **On Apps Script** — paste both files, leave `API_URL` and `API_TOKEN` empty
  at the top of `Index.html`, deploy as a web app. Google's own login is the
  lock.
- **On Cloudflare Pages** — add a `webAppToken` row to Settings with a long
  random value, redeploy the Apps Script web app with *Who has access: Anyone*,
  then fill in `API_URL` in your copy of `Index.html` and **leave `API_TOKEN`
  empty** so the token is never shipped in the file. You unlock once per device.

> The `/exec` URL plus the token is all anyone needs to reach your data. Keep
> both out of anywhere public, this repository included.

---

## Writing sequences

The Sequences sheet is **column pairs**. Each sequence is a message column with
an `awaitDays` column immediately to its left. The header of the message column
*is* the sequence name, and that exact text goes in a lead's `sequenceName`.

| A | B | C | D | E |
|---|---|---|---|---|
| `awaitDays` | `loomFollowUpMessage` | | `awaitDays` | `reEngage` |
| 3 | Step 1 text… | | 2 | Step 1 text… |
| 4 | Step 2 text… | | 3 | Step 2 text… |

`awaitDays` is the wait *after* that step goes out. Row 2 is step 1, row 3 is
step 2, and so on. The sequence ends when the next row is empty, and the lead is
marked `Done` the moment the last step sends.

### Variables

| Token | Becomes |
|---|---|
| `{{leadName}}` | column A of the lead's row |
| `{{firstName}}` | first word of that name |
| `{{leadEmail}}` | their address |
| `{{anything}}` | any column you add past `notes` — the header name is the token |

**A step is never sent with unresolved variables.** If a token in the message
is empty for that lead, or does not match any column, the send is held and the
lead is marked `Error` with a note naming the variable. Fill it in on the row
and set the status back to `Active`. Without this a blank `leadName` ships
"Hey , quick question" and a mistyped token ships `{{compnyName}}` straight to
the prospect.

Preview every new step with *Send test email* before it reaches anyone.

### Images

Put a Drive share link on its own line. The image is embedded inline in the
body, not attached.

```
Hey {{firstName}}, made you something:

[IMG 60%] https://drive.google.com/file/d/FILE_ID/view

Worth a look?
```

Size is optional: `[IMG]`, `[IMG 50%]`, `[IMG 300px]`. If the Drive fetch fails
the marker is stripped and the email goes out without it, rather than showing
the lead a broken placeholder.

---

## Settings

| Setting | Default | What it does |
|---|---|---|
| `followUpStartTime` | `08:00` | Window opens. Windows may cross midnight. |
| `followUpEndTime` | `18:00` | Window closes. The chain stops and resumes tomorrow. |
| `followUpSendingDays` | `Mon,Tue,Wed,Thu,Fri` | Three-letter days, comma separated. |
| `timezone` | `UTC` | **The clock everything runs on.** Set it to your *leads'* timezone, not yours. |
| `minDelayMinutes` | `4` | Random gap between real sends. |
| `maxDelayMinutes` | `8` | Leave these alone — see [Limits](#limits). |
| `maxSendsPerDay` | `50` | Hard stop for the day. |
| `replyCheckBeforeSend` | `TRUE` | **Keep this TRUE.** `FALSE` means you keep emailing people who already answered. |
| `resumeOnReply` *(per lead)* | `FALSE` | Tick on a lead's row to bring them back after they reply and go quiet. |
| `replyResumeDays` | `7` | Days of silence before a replied lead resumes. `0` is for testing only. |
| `replyResumeSequence` | — | Sequence a resumed lead moves onto. Blank means they carry on with the cold sequence they were in. |
| `oooAutoResume` | `FALSE` | Bring out-of-office leads back automatically. |
| `oooResumeDays` | `7` | How long to wait first. |
| `pauseAllFollowUps` | `FALSE` | The emergency brake. Stops everything, keeps all state. |
| `notificationEmail` | — | Where summaries and failure alerts go. Blank means silence. |
| `notificationFrequency` | `daily` | `daily` at the end of each chain, or `weekly`. |
| `notificationDay` | `Monday` | Weekly summary day. **Re-run *Create daily trigger* after changing this.** |
| `enrollDefaultSequence` | — | Sequence given to newly enrolled leads. Set it. |
| `enrollDefaultTotalSteps` | `0` | Step cap for new leads. `0` means no cap — the sequence ends when it runs out of rows. |
| `enrollDefaultResumeOnReply` | `FALSE` | Set `TRUE` and every new lead is opted into Track 2 automatically. |
| `webAppToken` | — | Password for the web app. Treat it like a key. |

### Timezone is the one people get wrong

The engine runs entirely on `timezone` — not your laptop, not the Apps Script
project. Set it to `America/New_York` and a 09:00–22:00 window means New York
business hours. That is correct if your leads are in New York and wrong if they
are not.

---

## Statuses

Only `Active` is ever picked up for sending.

| Status | Means | Comes back? |
|---|---|---|
| `Active` | In the sequence, waiting for `nextSendDate`. | — |
| `Replied` | They wrote back. Sequence stopped, `lastReplySnippet` holds their words. | Only with `resumeOnReply` |
| `OOO` | Auto-reply detected. Temporarily away. | Only with `oooAutoResume` |
| `Bounced` | Address is dead. Deliberately *not* `OOO` so nothing can revive it. | **Never** |
| `Done` | Sequence finished — last step sent, or the step cap hit. | Never |
| `Error` | Something failed. The `notes` column says what. | Fix it, set back to Active |
| `Paused` | You set this by hand. Nothing touches it. | When you say so |

### The notes column is load-bearing

Every pause, resume, bounce and failure is stamped with a date and appended,
never overwritten. Two of those entries are read back by the code:

- `Lead replied - paused.` — what `replyResumeDays` is counted from.
- `Resumed from Replied` — the cutoff that stops an old reply from re-pausing
  the lead forever.

**Do not edit or clear the notes on a paused lead.** Without the cutoff note a
resumed lead bounces straight back to `Replied` on the next run, and nothing
anywhere explains why.

---

## What to do when

### A lead replies

The system reads the whole thread before every send. It finds a message from
their address, stops, sets `Replied` with a dated note, and saves the first 120
characters of what they said. **You do nothing** — go have the conversation.

### They reply, then ghost you

With `resumeOnReply` ticked, after `replyResumeDays` of silence they go back to
`Active` and — if `replyResumeSequence` is set — move onto that sequence at step
1. Their old reply is marked handled so it cannot re-pause them.

To do it **now** instead of waiting: open the lead in the web app and use
**Resume sending**. Or run `resumeLeadNow` from the script editor.

> Never do this by editing the row. Setting `status` back to `Active` on its own
> does not work — the reply check re-finds their reply on the next run and
> reverts it. The button and `resumeLeadNow` write the cutoff note for you.

### You answer a lead by hand from Gmail

The system still stops. The check scans *every* message in the thread, not just
the last one, so your own reply sitting on top does not hide theirs. This is the
case that quietly breaks naive follow-up tools.

### An out-of-office bounces back

Recognised as `OOO` and noted, checked *before* the reply test so a robot never
gets filed as a human. Turn on `oooAutoResume` to bring them back after
`oooResumeDays`.

### The address is dead

Bounce wording is checked *first*, because mailer-daemon notices often contain
out-of-office phrasing too. Status goes to `Bounced`, permanently — auto-resume
can never touch it, so you will never loop on a dead mailbox.

### Gmail cuts you off mid-run

The chain stops for the rest of the day and emails you once. Every unsent lead
stays `Active` and is picked up tomorrow. Lower `maxSendsPerDay` — reputation
breaks long before the quota does.

### The chain stops and nothing is sending

If it died mid-run the lock clears itself after 30 minutes and the next
morning's run starts clean. If it could not create a trigger it emails you. Run
**Clean up stuck triggers**, then **Start follow-up chain now**.

### Moving a lead to a different campaign

Change `sequenceName`, set `sequenceStep` to `0` to start over, clear
`lastSentDate`, set `nextSendDate` to today, status `Active`. There is no hidden
state — the sequence and step are read fresh on every run.

### Re-testing a lead the same day

The system refuses to touch a lead it already handled today. That guard lives in
Script Properties, **not the sheet**, so editing the row will not clear it. Run
`unparkLead`.

### A lead is stuck on Error saying "variables unresolved"

A `{{token}}` in that step is empty for this lead, or names a column that does
not exist. The note says which. Fill the value in on the lead's row — or fix
the spelling in the sequence — then set the status back to `Active`.

This is why a newly enrolled lead needs its `leadName` and any custom columns
filled in before its first send. The system will hold rather than send a
half-written email, but it cannot invent the value.

### A lead went Done without ever being emailed

Its `sequenceName` is blank or misspelled. A blank name now raises a visible
`Error` with a note instead of silently finishing. Set `sequenceName` to match a
Sequences header exactly, reset `sequenceStep` to `0`, status back to `Active` —
then set `enrollDefaultSequence` so it cannot happen again.

---

## Limits

Your ceiling is not Gmail's quota. It is the delay that keeps you out of spam.

```
window   08:00 → 18:00   =  600 minutes
delay    4–8 min, avg 6
         600 ÷ 6          =  100 sends/day maximum
```

Skips are free — a parked, replied or bounced lead costs no time at all. Only
real sends spend the 4–8 minutes.

| Volume | Reality |
|---|---|
| 25/day | Effortless. The chain finishes by lunch. |
| 50/day | Half capacity. Comfortable — this is the sweet spot. |
| 100/day | Exactly at the limit, zero slack. One slow day spills into the next and compounds. Widen the window to 07:00–19:00 for 120/day. |

**Do not shorten the delay.** Fifty emails fired in one minute is the loudest
spam signal there is. The random 4–8 minute gap is what keeps you in the inbox.

Other ceilings: **20 triggers** per script (normal use is 4), **6 minutes** per
execution, and Gmail's own **500/day** on consumer accounts or **2,000** on
Workspace.

---

## Troubleshooting

Start in Apps Script → **Executions**. Every run logs what it decided and why.

### A healthy run

```
startDailyRun: pauseAllFollowUps=false
startDailyRun: isSendingDay=true | today=Sat
startDailyRun: isWithinSendingWindow=true | now=09:58 | window=09:00-22:00
Reply auto-resumed: lead@example.com (replied 2026-08-29, silent 7 days) - moved to "reEngage" at step 0
startDailyRun: found 1 lead(s) due today
processSingleLead: reply/OOO checks passed
processSingleLead: SEND SUCCESS | step=0 → 1 | nextSendDate=2026-08-30
```

### Diagnostic tools

| Run this | When |
|---|---|
| `debugReplyCheck` | A reply was not detected, or one was detected that should not have been. Prints the resolved account, the settings, the notes column, and every message in the thread with its timestamp and sender. |
| `unparkLead` | A lead will not be picked up and the log says `parked today`. |
| `resumeLeadNow` | Put one replied lead back into sending, safely, with the cutoff note. |
| `cleanUpStuckTriggers` | The chain died and will not restart. |

### Common log lines

| You see | It means |
|---|---|
| `skipped - parked today` | Already handled today. Expected — or run `unparkLead` to re-test. |
| `no leads due today` | Everything is future-dated, already sent, or not Active. Usually correct. |
| `a chain is already running` | The duplicate guard. Clears itself after 30 minutes. |
| `fetchTimeApiData ERROR … Address unavailable` | The time API is unreachable. Harmless — it falls back to the built-in clock, which gives the same answer. |
| `WARNING: replyCheckBeforeSend is FALSE` | You are about to email someone who already replied. Fix the setting. |
| `is not a column in the Sequences sheet` | `replyResumeSequence` does not match a header exactly. Case sensitive. |
| `CRITICAL: email sent but sheet write failed` | The lead got the email but the row did not update. The note gives the exact values to set by hand. |

**Check SendLog first.** Every attempt is recorded with a result — `Sent`,
`Failed`, `Skipped - …` or `Sent-WriteError`. It answers *"did this actually go
out?"* faster than anything else, and it is the one place that cannot be
confused by a row you edited afterwards.

---

## The files

| File | Where it goes |
|---|---|
| `Code.gs` | The engine. Paste into Apps Script bound to the spreadsheet. |
| `webapp/WebApp.gs` | Web app backend — `doGet` serves the page, `doPost` is the JSON endpoint with an explicit handler allow-list. |
| `webapp/Index.html` | The whole interface in one file. Must be named exactly `Index` in Apps Script. |

### The five sheets

| Sheet | Holds |
|---|---|
| `Settings` | Every setting, one per row. |
| `InboxScanner` | Scratch space for finding the right Gmail thread before enrolling. |
| `Sequences` | Your message templates, as `awaitDays` + message column pairs. |
| `ActiveFollowUps` | Every lead and its state. The heart of the system. |
| `SendLog` | One row per send attempt, with the result. |

### How the chain works

A single daily trigger starts it. From there the script schedules *itself*, one
lead at a time:

1. `startDailyRun` fires. If a chain is already alive it exits — two chains
   would double-send.
2. Pre-flight: pause flag, sending day, sending window, daily cap.
3. Out-of-office and replied leads are brought back if they are due.
4. If anything is due, it claims a lock and schedules the first
   `processOneLead`.
5. `processOneLead` deletes its own trigger, reads the Gmail thread, and either
   sends or parks the lead with a status.
6. **Only a real send costs time.** After one, it schedules the next 4–8 minutes
   out. A skip moves straight to the next lead, so a hundred parked leads do not
   eat the day.
7. The chain ends when nothing is due, the window closes, the cap is hit, or
   Gmail refuses. It releases the lock and emails you a summary if anything went
   out.
