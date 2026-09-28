# Earned Value Skill for Claude

Practical Claude skills for construction and infrastructure project controls.

## Skills

| Skill | What it does |
|---|---|
| **earned-value** | Works out PV, EV, AC, CPI, SPI, EAC and more. Builds CPR and WPM workbooks with live formulas and S-curves. Analyses your own cost reports and flags data traps. Explains why numbers look wrong. Writes variance commentary. Walks you through setting up EVM to AS 4817 / ISO 21508. |

## Needs

Python 3 with `openpyxl` (for the workbook and Excel analysis). Install it with `pip install openpyxl` if you don't have it.

## Install in Claude Code

Run these two commands inside Claude Code:

```
/plugin marketplace add Malrais/earned-value-skill
/plugin install earned-value@maisara-skills
```

To get updates later:

```
/plugin marketplace update maisara-skills
```

## Use in the Claude app (web, desktop or mobile)

Works on Free, Pro, Max, Team and Enterprise plans.

1. Turn on code execution: **Settings → Capabilities**. On Team or Enterprise plans, your admin controls this.
2. Download `earned-value.skill` from the [Releases](../../releases) page. If the upload won't accept it, rename it to `earned-value.zip`.
3. In Claude, go to **Customize → Skills**, click **+**, then **Create skill → Upload a skill**, and choose the file.
4. Ask something like *"Our CPI is 0.92 and SPI is 1.05. What does that mean?"*

## Feedback

Issues and suggestions are welcome. Open an issue on this repo.
