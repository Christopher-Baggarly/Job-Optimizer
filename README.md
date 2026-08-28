Automated Resume Engineering & Market Alignment Pipeline

The Business Problem
Career services and advising teams face a massive scaling bottleneck: manually reviewing, restructuring, and formatting client resumes, while simultaneously sourcing relevant local job market data. This manual process is prone to formatting errors, stylistic inconsistencies, and consumes hours of administrative overhead per client, limiting the number of individuals the organization can effectively serve.

The Architecture & Logic
This pipeline was engineered to completely remove the human formatting bottleneck by bridging legacy inputs with modern API data extraction and dynamic document rendering.

VBA to Python Bridge: Captures user parameters (target title, experience level) from a front-end interface and triggers the Python engine silently via command line.
Unstructured to Structured Data (OpenAI API): Ingests raw, unformatted .docx resumes and processes the text through a strict prompt structure. The engine evaluates job relevance on a 1-5 scale and forces the output into a strictly validated JSON schema.
Dynamic Templating Engine (Jinja2 / docxtpl): Bypasses standard Word manipulation libraries by using Jinja2 logic directly inside the .docx templates. The script maps the JSON payload to the template, dynamically generating horizontal/vertical skill matrices, formatting bullet points, and automatically spinning off an "Appendix" document if the client's job history exceeds the primary layout constraints.
Live Market Integration (Serper API): Automatically pings Google's search API based on the client's target title and locality, parses the JSON response, and generates a secondary Word document containing an actionable table of live, local job postings with direct application hyperlinks.
Technical Stack
Core: Python, Microsoft VBA
Libraries: docxtpl (Jinja2), python-docx, json, requests
External APIs: OpenAI (gpt-4o-mini), Serper (Google Search API)
The Impact
Reduced a multi-hour, highly manual advising and formatting process down to a seamless, zero-touch automated script. The system guarantees perfect formatting compliance across the organization, objectively grades historical job relevance, and arms the client with immediate, actionable labor-market data.

