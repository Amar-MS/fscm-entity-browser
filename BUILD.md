# How This Was Built

A summary of the approach — for anyone who wants to understand the process or build something similar.

---

## The Short Version

I am a business process consultant. I do not write code. This tool was built entirely by describing what I needed to **GitHub Copilot** in **VS Code**, testing the result, reporting what was wrong, and repeating. That's it.

---

## What You Need to Get Started

| Requirement | Where to get it |
|---|---|
| **VS Code** | https://code.visualstudio.com |
| **GitHub Copilot** | Install from VS Code Extensions — needs a Copilot subscription |
| **Microsoft Edge** | Already on Windows — used to load and test the extension |
| **A D365 F&SCM environment** | Any cloud environment you can log into |

---

## The Approach

### 1. Start with the problem, not the technology

I didn't start by thinking about code. I started with a frustration: in consulting, I constantly needed to understand what data lived inside a D365 entity — what fields existed, which were mandatory, what real data looked like. I described that frustration to Copilot as plainly as I would to a colleague, and asked it to build a browser extension.

### 2. One piece at a time

Rather than asking for everything at once, I asked for one feature at a time — entity list first, then schema view, then data query, then filters, then templates. Each step was small enough to test and verify before moving on. Copilot held the full context across the conversation and connected the pieces together.

### 3. Test, break, and describe what's wrong

Every time something didn't work, I described the symptom in plain language — "filters are returning an error", "the legal entity dropdown is empty", "column labels aren't showing". Copilot diagnosed the root cause and fixed it. I never needed to understand *why* the fix worked to move forward.

### 4. Raise requirements as you see them

Features like the Excel template, the M/T classification badges, and the column display names all came from me using the tool and thinking "I wish it did this." I raised them mid-conversation and they were added without starting over.

---

## What Copilot Handled

- All code — HTML, CSS, JavaScript, extension manifest
- XML parsing and OData query logic
- Bug diagnosis and root cause analysis
- Knowing D365-specific quirks (like how `$metadata` stores annotations, or why certain OData filters fail on non-string fields)

## What I Handled

- Knowing what the tool needed to do
- Testing each feature in a real D365 environment
- Spotting when something looked wrong or felt incomplete
- Deciding what to build next

---

## The Lesson

The gap between "I know what's needed" and "I can build it" has closed. The hardest part of building a useful tool is understanding the problem well enough to describe it clearly. That's a business skill, not a technical one.

If you want to build your own version — fork this repo, open VS Code with GitHub Copilot, and start describing what you need.

