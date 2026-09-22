# Dataroom vault

Public documents only. Every PDF is paired with a companion note that says where it
came from and when it was consulted — see `00_Admin/README.md`.

```
00_Admin/              register.csv (figures register), gaps.md, README.md
01_Financial/          10-K, URD, annual reports, analyst material
02_Permits_Regulatory/ permits, authorisations, inspection reports
03_Environment_HSE/    emissions, incidents, site environmental data
04_Litigation/         court filings, decisions, settlements
05_Commercial_Strategy/ tenders, contracts won, market notices
06_ESG/                ESG and sustainability reporting, ratings
07_Regulation_Texts/   EUR-Lex, Federal Register, EPA texts
Targets/<Name>.md      one note per target, linking to its documents
```

The MCP server only ever reads this folder. It never writes here, and it never
generates or summarises a number: it returns the passage, the page and the citation.
