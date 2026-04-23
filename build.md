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

---

## Version 1.2 — Number Sequence Health Analyzer

### Where the idea came from

In almost every D365 F&SCM implementation I have worked on, number sequences are set up once during go-live and never revisited. The default configuration often has no preallocation — or preallocation that was never tuned to the actual volume the system ended up handling. The result is silent performance drag that nobody connects to number sequences because it doesn't produce an obvious error.

I wanted a way to look at all number sequences in one place, immediately see which ones were likely causing problems, and get a concrete recommendation — not a vague "you should look at this."

### What I asked for

I described the idea to Copilot as: a tool inside the entity browser that shows all number sequences, flags any that have no preallocation or have the wrong configuration, and lets the user pick a volume tier — low, medium, high, very high — to get a recommendation matched to how busy their system actually is. I also wanted to be able to click into any sequence to see the detail and understand the specific issue.

### How it came together

The conversation covered three areas — what the data looked like in D365, what "good" and "bad" configuration meant in practice, and how to present it clearly without requiring the user to understand the underlying mechanics.

The volume tier idea came from a real consulting scenario: the right preallocation for a system processing 1,000 sales orders a month is completely different from one processing 1,000,000. Rather than showing a single recommendation, the tool lets the user set the context and then recalculates everything instantly.

D365 exposes number sequence configuration through the `SequenceV2Tables` OData entity. The tool calls this directly and falls back to scanning the environment's own entity list only as a safety net for non-standard environments.

### What it does

- Scans all number sequences in the environment across all companies
- Flags sequences with no preallocation, continuous sequences that are not optimised, and sequences that are running close to exhaustion
- Shows a progress bar for how much of each sequence's number range has been used, and estimates how long it will last at the selected volume
- Lets you filter by severity, company, or search by code
- Click any row to expand a full detail view with the specific issue and a plain-language recommendation

### What I handled

Knowing what good number sequence configuration looks like, describing the business scenarios that cause problems, testing the tool against a real environment, and identifying when the data being shown was wrong or incomplete.


