---
name: scopus-review
description: "Review and rewrite academic text using local Scopus and PDF evidence from PostgreSQL/pgvector. Use when the user invokes `$scopus-review`, asks to improve academic text with Scopus-backed scientific support, add ABNT author-date citations, validate claims against local Scopus/PDF collections, or produce a final text with ABNT references based on `scopus-search` results."
---

# Scopus Review

Use this skill to revise academic text by checking claims against the local `scopus-search` PostgreSQL/pgvector index and returning a clean, citation-backed final version.

Unless the user requests another language, write the final revised text in Brazilian Portuguese.

## Required Inputs

Require:

- an explicit base collection, named by the user as `collection`, `colecao`, `coleção`, or equivalent;
- the text to review.

There is no default collection.

If the collection is missing, run:

```powershell
python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py list-collections
```

Then list the available collections and ask which one to use.

If the text is missing, ask the user to provide the text to review.

## Collection Model

Each knowledge base usually has two related collections:

- `<collection>`: Scopus metadata, abstracts, keywords and bibliographic records.
- `<collection>_pdf`: PDF chunks/full-text evidence.

Use them with different roles:

- `<collection>` identifies candidate studies, metadata, DOI, title, authors, source title, year and abstract-level support.
- `<collection>_pdf` provides stronger textual support from full-text chunks when a claim needs deeper validation.

Do not treat the two collections as interchangeable.

## Execution Mode

Prefer keeping the main thread as an orchestrator.

For long, complex, or high-stakes reviews, delegate the review workflow to a strong reasoning subagent when available.

The review worker must perform:

- claim decomposition;
- query planning;
- compact evidence retrieval;
- evidence judgment;
- final rewrite;
- ABNT citation and reference formatting.

If subagent execution is unavailable in the current runtime, perform the same workflow in the main thread. Do not fail only because a specific model, reasoning effort, tool, or subagent runtime is unavailable.

Preserve any explicit model, reasoning effort, skill, or specialized agent configuration already selected by the user or runtime.

## Token Budget Policy

Minimize tool calls and returned context.

Default behavior:

- create at most 2 English semantic queries per claim;
- use `--top-k 3`;
- search `<collection>` first;
- search `<collection>_pdf` only when needed;
- do not search both collections for every query by default;
- deduplicate immediately after each search batch;
- retain only compact evidence notes, not raw chunks.

Use `--top-k 5` only when:

- diagnostic mode is requested;
- the first pass finds no direct support;
- the claim is central to the user's argument and evidence remains weak after the first pass.

Do not keep or repeat raw chunks, long metadata dumps, scores, or search logs unless the user explicitly asks for diagnostic mode.

## Retrieval Constraints

Use only the deterministic local wrapper.

Do not use:

- web search;
- browser tools;
- external MCPs;
- model memory;
- general internet sources;
- uncited prior knowledge;

as evidence for the revised text.

Search commands must follow this contract:

```powershell
python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py list-collections

python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py search --collection <collection> --query "<english query>" --original-query "<claim text>" --top-k 3
```

Before searching, run `list-collections` once and determine which of these exist:

- `<collection>`;
- `<collection>_pdf`.

Do not search unrelated collections.

## Dual Collection Search Strategy

Use Scopus metadata first and PDF chunks only when needed.

Default search order:

1. Search `<collection>` first with `--top-k 3`.
   - Use it to identify candidate studies.
   - Use it to recover bibliographic metadata.
   - Use it to assess abstract-level support.
   - Prefer it for reference construction.

2. Search `<collection>_pdf` only when:
   - `<collection>` returns fewer than 2 directly relevant candidate sources;
   - Scopus-level evidence is too generic;
   - the claim needs full-text support;
   - the paragraph contains a specific methodological, empirical, regulatory, technical or result-oriented statement;
   - the claim is important for the user's argument and requires stronger validation;
   - the user explicitly asks for deeper evidence or PDF validation.

3. If `<collection>` returns no useful candidates, search `<collection>_pdf` directly.

4. If `<collection>_pdf` does not exist, continue with `<collection>` only.

5. When both collections return the same work:
   - merge evidence by DOI first;
   - if DOI is absent, merge by normalized title;
   - use Scopus metadata for the reference;
   - use PDF chunks for stronger textual support.

6. Do not search both collections for every query by default.

7. Avoid redundant searches with nearly identical queries.

## Review Workflow

1. Split the input into claim-sized units.
   - A claim can be one sentence, two closely related sentences, or one short paragraph.
   - Keep claims narrow enough that a citation can logically support each one.

2. For each claim, create at most 2 English semantic queries.
   - Preserve the exact intent of the claim.
   - Do not add unrelated theory, methods, or keywords only to improve retrieval.
   - Prefer queries that preserve the domain, method, population, object, and outcome of the claim.

3. Retrieve evidence using the dual collection strategy.
   - Start with `<collection>`.
   - Use `--top-k 3` by default.
   - Search `<collection>_pdf` only when Scopus evidence is absent, weak, generic or insufficient.
   - Avoid redundant searches.

4. Merge and deduplicate retrieved candidates.
   - Deduplicate by DOI first.
   - If DOI is absent, deduplicate by normalized title.
   - Prefer metadata DOI over DOI inferred from snippets.
   - Avoid duplicate references when the same work appears in both `<collection>` and `<collection>_pdf`.

5. Keep a compact evidence cache for each claim.
   - Keep DOI, title, authors, year, source title, and one short evidence note.
   - Keep support judgment.
   - Do not retain raw chunks unless diagnostic mode is requested.

6. Judge retrieved evidence before citing.
   - Use a study only when its title, metadata, abstract or retrieved snippet directly supports the claim's meaning.
   - Treat vague topical similarity as insufficient support.
   - Prefer stronger and more direct evidence over higher-score but generic chunks.
   - Do not cite a source if the retrieved evidence does not support the sentence or paragraph.

7. Rewrite the text.
   - Keep supported claims and cite them where the support applies.
   - Preserve the author's original meaning when the literature supports it.
   - Improve academic clarity, cohesion, precision, and caution.
   - Do not add unsupported claims.
   - Do not strengthen claims beyond the recovered evidence.

## Handling Weak or Missing Evidence

When a claim lacks sufficient evidence:

- do not leave a strong unsupported claim without citation;
- rewrite it in a narrower, cautious, conditional, or contextual form when possible;
- phrase it as a research motivation, hypothesis, limitation, or possible implication when direct support is weak;
- if no evidence is found and the claim cannot be safely softened, remove the factual assertion while preserving the author's broader intent;
- do not invent references;
- do not cite sources based only on topical similarity.

Use cautious wording when needed, such as:

- "pode contribuir";
- "pode favorecer";
- "sugere-se";
- "em determinados contextos";
- "em aplicações especializadas";
- "quando apoiado por corpus específico";
- "a literatura recuperada indica".

## Citation Rules

Use ABNT author-date citations in the text.

Examples:

```text
(SILVA; PEREIRA, 2024)
(SILVA et al., 2024)
```

Rules:

- Cite at the sentence or paragraph where the evidence applies.
- For paragraphs up to 80 words, place at least one or two citations at the end of the paragraph that support the paragraph as a whole.
- For paragraphs above 80 words, place at least one citation in the first half and at least one citation at the end.
- You may adjust paragraph wording to make the cited evidence support the full paragraph more coherently.
- If the cited author string cannot be parsed reliably, use the first clear responsible name from metadata.
- If no responsible name can be parsed, do not cite that item.
- Do not cite a source if the retrieved evidence does not support the sentence.
- Do not use uncited factual claims unless they are purely transitional, methodological, or already contained in the user's text in a cautious form.

## Reference Rules

Build the final references from retrieved metadata.

Include, when available:

- authors;
- title;
- source title;
- year;
- DOI.

Use a consistent ABNT-like reference format:

```text
SOBRENOME, Prenomes; SOBRENOME, Prenomes. Título do trabalho. Nome da fonte, ano. DOI: ...
```

Reference rules:

- Include every cited work.
- Include only cited works.
- Deduplicate references by DOI first, then by normalized title.
- Prefer complete metadata from the retrieved result.
- Prefer Scopus metadata over PDF chunk metadata when both refer to the same work.
- Do not invent missing metadata.
- If DOI is unavailable, omit DOI rather than fabricating one.
- Keep references in alphabetical order by first author surname when possible.

## Output Format

Default mode is final-output mode.

Return only:

1. the improved text, with ABNT author-date citations placed in the related sentence or paragraph;
2. exactly two blank lines;
3. the heading `Referências bibliográficas`;
4. ABNT-style references for every cited work, and only cited works.

Do not include:

- audit matrix;
- query list;
- raw chunks;
- scores;
- process notes;
- evidence tables;
- internal reasoning;

unless the user explicitly asks for audit, diagnostics, debug, matrix, or evidence report.

## Diagnostic Mode

If the user explicitly asks for `auditoria`, `diagnóstico`, `debug`, `matriz`, `evidence report`, or equivalent, return a review matrix instead of final-output mode.

In diagnostic mode only:

- use up to 3 English queries per claim;
- use `--top-k 5`;
- search both `<collection>` and `<collection>_pdf` when available;
- include concise evidence details but avoid dumping full chunks unless requested.

The matrix must include:

- original claim;
- generated English queries;
- searched collections;
- selected evidence;
- support judgment;
- revised wording;
- cited reference.

Use support judgments:

```text
supported
partially supported
weak support
unsupported
```

Do not expose raw internal reasoning. Provide concise justification based on retrieved evidence.

## Quality Checks Before Final Answer

Before returning the final deliverable, verify:

- every citation in the text appears in the references;
- every reference is cited in the text;
- no source is cited only because of vague topical similarity;
- unsupported claims were softened, reframed, or removed;
- citations are placed close to the claims they support;
- duplicate references were removed;
- Scopus metadata and PDF evidence were merged when they refer to the same work;
- the final text preserves the user's original meaning as much as the evidence allows;
- the output follows the required format exactly.

## Recommended Delegation Prompt

When delegating to a subagent, use this structure:

```text
Use $scopus-review to revise the text below.

Collection: <collection>
Collection strategy:
- Search <collection> first for metadata, abstracts and candidate studies.
- Search <collection>_pdf only if Scopus-level evidence is weak, missing, generic or if the claim needs full-text validation.

Workflow:
- split the text into claim-sized units;
- generate at most 2 English queries for each claim;
- search only with:
  python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py
- use --top-k 3 by default;
- avoid redundant searches;
- do not search both collections for every query by default;
- judge support before citing;
- deduplicate sources by DOI, then normalized title;
- prefer Scopus metadata for references and PDF chunks for stronger textual support;
- keep only compact evidence notes;
- soften, narrow, or reframe unsupported claims;
- do not invent references;
- use ABNT author-date citations;
- return only the final revised text, exactly two blank lines, and Referências bibliográficas.

Text:
<text>
```
