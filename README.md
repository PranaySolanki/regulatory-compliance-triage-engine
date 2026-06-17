# On-Premise Regulatory Compliance Triage Engine

An enterprise-grade, localized automation hub built in VBA designed to parse, analyze, and structure high-priority regulatory advisories (RBI, CERT-In) from Microsoft Outlook directly into a structured GRC data tracking dashboard.

## 🏢 The Problem Statement
Financial institutions receive critical cyber threat intelligence and compliance directives daily. Manual sorting, text extraction, and deduplication consume hours of analyst capacity. Crucially, strict financial data-privacy standards prohibit routing these sensitive infrastructure advisories through public cloud AI APIs (such as OpenAI or Claude) for automated parsing. 

## 🛡️ The Solution
This engine serves as a **Zero-Cloud Dependency Triage Solution**. Running entirely within the local application memory sandbox, it securely hooks into desktop Outlook via the Object Model, implements a stateful deduplication filter, isolates unstructured information boundaries using an alternative keyword matrix, and scales a functional monitoring hub in seconds.

### Core Features
* **Dynamic Data Pipeline:** Automatically extracts and structures critical routing metadata, including the transmission date, sender identity, and internal target mailbox.
* **Proactive Attachment Flagging:** Seamlessly evaluates emails for files, applying a clear structural indicator so analysts instantly recognize when an advisory requires physical document review inside Outlook.
* **Shifting-Text Boundary Parser:** Leverages an algorithmic array structure to cleanly isolate multi-line text blocks across varying regulatory writing styles, mapping threat overviews and technical recommendations dynamically.
* **Stateful Long-Term Archive:** Feeds a dedicated structural history log that preserves a permanent audit trail, ensuring compliance records remain intact even after the active tracking dashboard is cleared.
* **Inline Deduplication Engine:** Automatically cross-references previously processed subjects within the long-term log to skip redundant entries, optimizing workspace efficiency.

## 🖥️ How it Works (System Architecture)
1. **Target Timeline Ingestion:** Reads an analyst-specified timeline day filter directly from the interface dashboard.
2. **Inbox Stream Scanning:** Re-sorts and evaluates incoming email streams within Outlook.
3. **Property & Context Extraction:** Pulls core transmission paths while evaluating attachment presence to populate conditional formatting triggers.
4. **Symmetrical Dual-Tab Population:** Inserts active records at the top chronologically across both active working screens and permanent logs to maintain a seamless audit trail.

## 📋 How to Use
1. **Configure Timeline Filter:** Open the main `Dashboard` tab and input the desired historical lookback window (in days) into cell `B4` (e.g., entering `10` will scan emails received over the past 10 days).
2. **Execute Sync Engine:** Click the green **Fetch Advisories** button on the dashboard interface. This triggers the localized Outlook API hook.
3. **Review Extracted Data:** The engine clears the active dashboard view and repopulates it with live entries matching the keywords (`RBI`, `CERT`, etc.), complete with extracted overviews and structural attachment indicators.
4. **Audit Long-Term History:** Navigate to the `Compliance_Archive` tab to review the deduplicated chronological record ledger of all captured historical alerts.




https://github.com/user-attachments/assets/ce693200-0f24-494e-bcfb-a1f6af3c1c1f



