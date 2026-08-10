# Field report - the first real task split across both servers, 2026-08-10

The owner asked, in an ordinary working session: *"using the new exceltask mcp
can you fix this macro?"* with a real workbook attached. The agent used the
original Excel MCP for part of the task and ExcelTask for the rest. Its own
account, relayed by the owner, verbatim in the parts that matter:

> "They're not competitors - I needed both. excel-mcp did the thinking work.
> `vba list` told me the module and procedure names existed at all; `vba view`,
> `worksheet list`, and a dozen range get-values calls let me compare the
> macro's hard-coded rows against real sheet layouts. Without that I'd never
> have found the off-by-one. excel-task can't discover anything - it demands the
> component and procedure name upfront, so on its own it's blind."

> "excel-task did the dangerous part properly. Refused to save in place - your
> original file was never at risk. Required a SHA-256 of the existing code - if
> the file had changed under me, the edit would abort instead of clobbering.
> Reopened the saved file and re-checked the fingerprint - proof it actually
> worked, not just 'no error'. Notably, excel-mcp has `vba update`, which would
> have made the same edit in one call - but with none of those three
> protections."

> "Verdict: excel-mcp for reading and exploring, excel-task for any edit you'd
> be upset to get wrong. If excel-task added a read-only 'list modules and
> procedures' action, it'd stand alone."

## What this settles

**The discovery gap is confirmed as the split point.** The task went to the
other server exactly where ExcelTask offers nothing: learning what a workbook
contains. The fix - the audit listing macro components and procedures, with the
schema routing unknown names to it - was built the same day and held for this
confirmation before release.

**The safety design was validated by its user.** The three guardrails were not
tolerated; they were named as the reason to prefer ExcelTask for any edit that
matters, against a one-call alternative with none of them.

**Two defects from real use:**

1. The macro policy rejected the caller twice, teaching one rule per round trip.
   Fixed: every unmet requirement now arrives in a single rejection, held by a
   test.
2. Whole-procedure replacement cannot introduce a module-level constant; the
   agent worked around it by scoping the constant inside the procedure. Recorded
   as a roadmap candidate, not built - one occurrence with a clean workaround is
   not yet demand.

## What this deliberately leaves open

Name discovery was not the whole story: finding the actual bug took **reading
cell values** across sheets, which ExcelTask does not do and was designed not to
do - its receipts carry no workbook data. So even with macro discovery, "fix
this macro by comparing it to the sheets" remains a two-server task.

That is the next strategic question, and it should be answered by more real use,
not by this one incident: either the two-server split is the intended end state
- the original for reading and exploring, ExcelTask for edits that must not go
wrong - or ExcelTask grows a bounded, read-only inspection of values. The
second would be a significant widening of what the product promises never to
return, and it is not being taken lightly or soon.
