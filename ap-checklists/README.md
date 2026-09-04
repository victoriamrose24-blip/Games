# AP Process Checklists — Skeleton

Scope: **review, schedule, pay.** Vouching/entry is done upstream by someone else.

Everything in `[SQUARE BRACKETS]` is a placeholder for you to replace with how it
actually works at your shop. Delete any step that doesn't apply. Add the ones I
couldn't know about.

## How to use these

Five cards, each a separate file. You never run more than one at a time.

| Card | File | Style | Frequency |
|---|---|---|---|
| A. Invoice review gate | `A-invoice-review.md` | DO-CONFIRM | Per invoice, many/day |
| B. Run build | `B-run-build.md` | READ-DO | Per run |
| C. Release / transmit | `C-release.md` | READ-DO | Per run |
| D. Post-run | `D-post-run.md` | READ-DO | Per run |
| E. Vendor bank / remit-to change | `E-vendor-bank-change.md` | READ-DO | On trigger only |

**DO-CONFIRM** — do the work in your normal flow, then stop at the pause point and
confirm the short list. Used where volume is high and reading every line would get
abandoned by Thursday.

**READ-DO** — read the step, then do it. Used where the task is rare, or where the
error is expensive and hard to reverse.

## The Never-Skip tier

These five get done even on the worst day. Everything else on every card is hygiene
and can flex under pressure. If you find yourself skipping one of these, that's the
signal to stop and slow down, not to keep going.

1. Duplicate check
2. Approval present, and from someone with authority for that amount
3. Remit-to on the invoice matches the vendor master
4. Batch total and count reconciled at release
5. Funding account correct

`[EDIT THIS LIST once you've listed your actual past misses. It should reflect what
has really gone wrong for you, not what sounds important.]`

## The interruption rule

An item's status lives in the system, never in your head.

Every invoice sits in exactly one named state at all times:

`[Received] -> [Reviewed] -> [Scheduled] -> [Released] -> [Paid]`
plus the parking states: `[On hold]`, `[In dispute]`, `[Pending vendor verification]`

If you get pulled away mid-item, you park it in the state it is actually in —
including a "started, not finished" state — **before** you look at anything else.
When you come back you never ask "where was I." You ask the queue.

`[EDIT: replace the states above with the actual status values available in
[YOUR SYSTEM]. If your system has no usable status field, use a spreadsheet or a
saved view — but it has to live outside your head.]`

## The re-transmit rule

**Never re-send a payment file based on your own recollection.** If you can't tell
whether a run went out, the bank's confirmation is the only source of truth. Go look
at bank activity before you touch anything.

This is the single most expensive interruption failure in this job. It gets its own
rule because it deserves one.

## Trigger index

Rare and conditional steps fail because you forget they exist, not because you don't
know how to do them. So the triggers live inline in the cards, and the procedures
live in their own files. You only have to remember the trigger.

| If you see this... | Stop and run... |
|---|---|
| New vendor, or changed remit-to / bank detail | Card E |
| Invoice is on hold or in dispute | `[YOUR HOLD PROCEDURE]` |
| Open credit memo for this vendor | `[YOUR CREDIT MEMO PROCEDURE]` |
| Amount over `[$THRESHOLD]` | `[YOUR ADDITIONAL APPROVAL PROCEDURE]` |
| Payment returned or rejected by bank | `[YOUR RETURNS PROCEDURE]` |
| Vendor requests urgent/off-cycle payment | `[YOUR OFF-CYCLE PROCEDURE]` |

`[ADD YOUR OWN. Every time you hit a situation the cards don't cover, add a row
here rather than handling it from memory.]`

## Maintaining this

Once a month, or after any miss:

- What went wrong? Which card should have caught it?
- Was the step missing from the card, or was it on the card and skipped?
  - Missing -> add it
  - Skipped -> ask whether the card is too long, or whether the step belongs in
    the Never-Skip tier
- Is any card long enough that you've stopped actually running it? Cut it.

A checklist you stopped using is worse than no checklist, because you think you're
covered. Length is the usual cause.
