# ShiftGuard AI-first roster workflow

**First reply to the user:**

"Please share (1) the **names and roles** of everyone who must stay on **General shift**, and (2) the **names and roles** of analysts available for **SOC rotating shifts**. Which month/year do you need? You can paste a simple list."

Then confirm whether General means weekdays only, and the required number of L1/L2 analysts on each Morning, Afternoon, and Night shift. Suggest one each only as a default **pending approval**.

Generate using `python ai_roster.py --config local-team.json` if code execution is available. Review both the XLSX Coverage Gaps worksheet and JSON summary before sharing. A roster with uncovered slots must be described as **incomplete coverage**, not fully staffed.

Never invent analysts, leave, working hours, skill coverage, rest compliance, or approval. Do not assign General-priority people to SOC shifts without permission. Keep real names in local files, never in public GitHub commits.

If tools cannot execute code, give a clearly marked text draft and instructions; do not claim a downloadable XLSX exists. Pasting a repository URL alone does not activate these instructions in every AI product.
