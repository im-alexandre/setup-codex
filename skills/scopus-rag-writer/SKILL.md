---
name: scopus-rag-writer
description: "Write academic text from scratch using local Scopus and PDF evidence from PostgreSQL/pgvector. Use when the user invokes `$scopus-rag-writer`, asks to write a paragraph, section, introduction, theoretical background, discussion, or article fragment grounded in a local Scopus/PDF collection with ABNT author-date citations."
---

# Scopus RAG Writer

Use this skill to write academic text grounded in a local `scopus-search` PostgreSQL/pgvector index.

This skill is for **generation from evidence**, not only revision. It should retrieve evidence first, classify the evidence, synthesize the retrieved literature, and then write the requested text with ABNT author-date citations and references.

Unless the user requests another language, write in Brazilian Portuguese.

## Required Inputs

Require:

- an explicit base collection, named by the user as `collection`, `colecao`, `coleção`, or equivalent;
- a writing task, such as paragraph, section, introduction, theoretical background, discussion, literature review, justification, objective, or subsection;
- the topic, problem, claim, or outline to write about.

There is no default collection.

If the collection is missing, run:

```powershell
python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py list-collections
```

Then list the available collections and ask which one to use.

If the writing task is missing or underspecified, ask for the missing scope only when necessary.

## Collection Model

Each knowledge base usually has two related collections:

- `<collection>`: Scopus metadata, abstracts, keywords and bibliographic records.
- `<collection>_pdf`: PDF chunks/full-text evidence.

Use them with different roles:

- `<collection>` identifies candidate studies, metadata, DOI, title, authors, source title, year, keywords and abstract-level support.
- `<collection>_pdf` provides stronger textual support from full-text chunks when the text needs methodological, empirical, conceptual, regulatory, or result-oriented detail.

Do not treat the two collections as interchangeable.

## Execution Mode

Prefer keeping the main thread as an orchestrator.

For long, complex, or high-stakes writing tasks, delegate the evidence search and drafting workflow to a strong reasoning subagent when available.

The writing worker must perform:

- writing intent decomposition;
- query planning;
- compact evidence retrieval;
- evidence classification;
- evidence synthesis;
- academic drafting;
- ABNT citation and reference formatting.

If subagent execution is unavailable in the current runtime, perform the same workflow in the main thread. Do not fail only because a specific model, reasoning effort, tool, or subagent runtime is unavailable.

Preserve any explicit model, reasoning effort, skill, or specialized agent configuration already selected by the user or runtime.

## Retrieval Constraints

Use only the deterministic local wrapper.

Do not use:

- web search;
- browser tools;
- external MCPs;
- model memory;
- general internet sources;
- uncited prior knowledge;

as evidence for the generated text.

Search commands must follow this contract:

```powershell
python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py list-collections

python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py search --collection <collection> --query "<english query>" --original-query "<writing intent or claim>" --top-k 5
```

Before searching, run `list-collections` once and determine which of these exist:

- `<collection>`;
- `<collection>_pdf`.

Do not search unrelated collections.

## Retrieval Budget Policy

Balance evidence quality and context cost.

Default behavior:

- create 3 to 5 English semantic queries for the writing task;
- use `--top-k 5` for `<collection>`;
- search `<collection>` first;
- search `<collection>_pdf` when Scopus-level evidence is weak, generic, insufficient, or when the requested text needs deeper validation;
- use `--top-k 5` for `<collection>_pdf` when searched;
- deduplicate immediately after each search batch;
- retain only compact evidence notes, not raw chunks.

For central arguments, literature review paragraphs, theoretical framing, or weak first-pass retrieval:

- create up to 6 English semantic queries;
- use `--top-k 8`;
- include at least one literal query based on the user's wording;
- include at least one domain-specific query;
- search both `<collection>` and `<collection>_pdf` if first-pass evidence is vague or insufficient.

Diagnostic mode:

- create up to 8 English semantic queries;
- use `--top-k 10`;
- search both `<collection>` and `<collection>_pdf`;
- include concise evidence details but avoid dumping full chunks unless requested.

## Dual Collection Search Strategy

Use Scopus metadata first and PDF chunks when needed.

Default search order:

1. Search `<collection>` first with `--top-k 5`.
   - Use it to identify candidate studies.
   - Use it to recover bibliographic metadata.
   - Use it to map the literature.
   - Prefer it for reference construction.

2. Search `<collection>_pdf` with `--top-k 5` when:
   - `<collection>` returns fewer than 3 directly relevant candidate sources;
   - Scopus-level evidence is too generic;
   - the requested text needs full-text support;
   - the user asks for methodological, empirical, regulatory, technical, conceptual, or result-oriented content;
   - the user asks for a literature review, theoretical background, discussion, or critical comparison;
   - the user explicitly asks for deeper evidence or PDF validation.

3. If evidence remains weak after the first pass:
   - run a second pass with up to 6 queries;
   - use `--top-k 8`;
   - include one literal query based on the user's wording;
   - include one query with explicit domain terms.

4. If `<collection>` returns no useful candidates, search `<collection>_pdf` directly.

5. If `<collection>_pdf` does not exist, continue with `<collection>` only.

6. When both collections return the same work:
   - merge evidence by DOI first;
   - if DOI is absent, merge by normalized title;
   - use Scopus metadata for the reference;
   - use PDF chunks for stronger textual support.

7. Avoid redundant searches with nearly identical queries.

## Evidence Classification Gate

Classify every candidate source before citing it.

Use these labels:

- `core`: directly addresses the concept, method, mechanism, or empirical claim being written.
- `applied`: applies the relevant concept in a specific domain and can support examples, applications, or domain-specific claims.
- `adjacent`: shares broad vocabulary or a related theme but does not directly support the claim.
- `reject`: does not support the claim.

Citation rules:

- Use `core` sources for conceptual, methodological, or general claims.
- Use `applied` sources only for examples, empirical illustrations, or domain-specific claims.
- Do not use `adjacent` sources for conceptual claims.
- Never cite `reject` sources.
- If only `adjacent` or `reject` sources are found, run a second-pass retrieval with more literal and domain-specific queries.
- If second-pass retrieval is still weak, write a narrower cautious statement or state that the local collection did not provide direct evidence.

For conceptual claims about RAG, embeddings, LLMs, hallucination, grounding, traceability, provenance, auditability, source attribution, evidence-based generation, semantic search, dense retrieval, or vector databases:

- prefer methodological papers, surveys, benchmarks, evaluation papers, or papers directly addressing those concepts;
- use domain-applied papers only as examples of use, not as support for the general concept;
- reject papers that discuss only generic evaluation, ranking, prioritization, risk, governance, education, or decision-making without addressing the relevant RAG/embedding/LLM concept.

## Rejection Rules

Reject sources that are only generically related.

Do not cite a source merely because it discusses:

- ranking;
- prioritization;
- evaluation;
- decision-making;
- matrices;
- assessment;
- education;
- risk;
- governance;
- process improvement;
- generic information systems;

unless it directly supports the specific text being generated.

For claims about RAG, embeddings, LLMs, hallucination, grounding, traceability, provenance, auditability, source attribution, evidence-based generation, semantic search, dense retrieval, or vector databases, prefer sources that explicitly discuss one or more of those concepts.

If no directly relevant source is found, write a narrower, cautious statement or state that the local collection did not provide sufficient evidence.

## Writing Workflow

1. Identify the requested output type.
   - paragraph;
   - section;
   - subsection;
   - introduction;
   - theoretical background;
   - literature review;
   - discussion;
   - justification;
   - research problem;
   - objective;
   - academic synthesis.

2. Decompose the writing task into 2 to 5 evidence needs.
   - concept definition;
   - empirical support;
   - method/application;
   - limitation;
   - comparison;
   - implication;
   - domain-specific example.

3. Generate English semantic queries for each evidence need.
   - Preserve the user's topic and intent.
   - Do not add unrelated theory or methods only to improve retrieval.
   - Include one literal query when the user provides a specific claim.

4. Retrieve evidence using the dual collection strategy.

5. Merge and deduplicate retrieved candidates.
   - Deduplicate by DOI first.
   - If DOI is absent, deduplicate by normalized title.
   - Prefer metadata DOI over DOI inferred from snippets.

6. Classify the evidence before drafting.
   - Label each candidate as `core`, `applied`, `adjacent`, or `reject`.
   - Keep `core` and suitable `applied` sources.
   - Do not cite `adjacent` or `reject` sources.
   - Run second-pass retrieval when only weak candidates are found.

7. Build a compact evidence plan before drafting.
   - Keep DOI, title, authors, year, source title, source type label, and one short evidence note.
   - Group evidence by function: definition, support, contrast, limitation, example, implication.
   - Do not retain raw chunks unless diagnostic mode is requested.

8. Draft the requested text.
   - Write academically and coherently.
   - Synthesize sources rather than listing them.
   - Place citations close to the claim they support.
   - Use cautious wording when the evidence is partial.
   - Do not add unsupported factual claims.
   - Do not overstate what the retrieved evidence supports.

9. Build the final references from cited metadata only.

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
- Do not use uncited factual claims unless they are purely transitional or explicitly framed as the user's proposed argument.

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

1. the generated academic text, with ABNT author-date citations placed in the related sentence or paragraph;
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

If the user explicitly asks for `auditoria`, `diagnóstico`, `debug`, `matriz`, `evidence report`, or equivalent, return a writing evidence matrix instead of final-output mode.

In diagnostic mode only:

- use up to 8 English queries;
- use `--top-k 10`;
- search both `<collection>` and `<collection>_pdf` when available;
- include concise evidence details but avoid dumping full chunks unless requested.

The matrix must include:

- writing need or claim;
- generated English queries;
- searched collections;
- selected evidence;
- evidence classification (`core`, `applied`, `adjacent`, `reject`);
- support judgment;
- proposed wording;
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
- conceptual claims are supported by `core` sources whenever possible;
- applied sources are used only for applied/domain-specific claims;
- unsupported claims were softened, reframed, or removed;
- citations are placed close to the claims they support;
- duplicate references were removed;
- Scopus metadata and PDF evidence were merged when they refer to the same work;
- the final text answers the requested writing task;
- the output follows the required format exactly.

## Recommended Delegation Prompt

When delegating to a subagent, use this structure:

```text
Use $scopus-rag-writer to write the requested academic text.

Collection: <collection>
Writing task: <paragraph/section/introduction/theoretical background/etc.>
Topic or outline:
<topic>

Collection strategy:
- Search <collection> first for metadata, abstracts and candidate studies.
- Search <collection>_pdf if Scopus-level evidence is weak, missing, generic or if the writing task needs full-text validation.

Workflow:
- decompose the writing task into evidence needs;
- generate 3 to 5 English queries by default;
- for central or weakly supported arguments, generate up to 6 queries, including one literal query;
- search only with:
  python C:\Users\imale\.codex\skills\scopus-search\scripts\scopus_search.py
- use --top-k 5 by default;
- use --top-k 8 for second-pass retrieval when evidence is weak;
- avoid redundant searches;
- classify every candidate source as core, applied, adjacent, or reject;
- reject sources that are only vaguely related;
- use core sources for conceptual claims;
- use applied sources only for examples or domain-specific claims;
- deduplicate sources by DOI, then normalized title;
- prefer Scopus metadata for references and PDF chunks for stronger textual support;
- keep only compact evidence notes;
- synthesize sources into academic prose;
- do not invent references;
- use ABNT author-date citations;
- return only the final generated text, exactly two blank lines, and Referências bibliográficas.
```
