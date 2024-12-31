Automated Lead Capture from Keyloop Emails via Outlook (Mac/Windows)
====================================================================

> **Project Type**: Freelance Automation\
> **Technologies**: Node.js, Puppeteer, Outlook Web (PWA)\
> **Platform**: Mac & Windows

* * * * *

Project Overview
----------------

This project automates the process of capturing automotive sales leads from **Keyloop** emails in **Outlook**. Whenever a new "reply by email" lead arrives, the script instantly clicks the unique link---often within one second---thereby assigning the lead to the user. By reacting faster than competitors, the salesperson greatly increases their chance of winning the sale.

**Why It Matters**

-   **Speed to Lead**: Keyloop assigns the lead to whoever clicks first.
-   **Local-Only Execution**: Needed to bypass restricted dealership PCs.
-   **Cross-Platform**: Delivered for both Mac and Windows.

* * * * *

Key Features
------------

1.  **Ultra-Fast Detection**

    -   Monitors Outlook's web inbox in real time, grabbing new leads within ~1 second of arrival.
2.  **Simple Start/Stop**

    -   Script-based approach that runs locally without requiring special permission on corporate computers.
3.  **Targeted Email Filtering**

    -   Listens for a specific sender (`noreply@leadmanager.com`) and timestamps to ensure only fresh leads are captured.
4.  **Partial Auto-Refresh**

    -   Minimizes Outlook stalls and ensures the script continues running smoothly over extended periods.

* * * * *

Technical Overview
------------------

-   **Node.js & Puppeteer**: Automates Google Chrome to interact with Outlook's web UI.
-   **MutationObserver**: Detects new emails in real time, using DOM observation.
-   **Timestamp Validation**: Ensures only leads that arrive after script launch are clicked, avoiding older messages.
-   **Cross-Platform Setup**:
    -   **Mac**: Shell script (`.sh`) handles dependencies and permissions.
    -   **Windows**: `.bat` file plus Node.js installation for straightforward execution.

* * * * *

Challenges & Resolutions
------------------------

1.  **Corporate Environment Restrictions**

    -   *Challenge*: Couldn't install software on dealership PCs.
    -   *Solution*: Created a self-contained script for personal devices.
2.  **DOM Selectors**

    -   *Challenge*: Outlook's UI changes frequently, causing script to fail.
    -   *Solution*: Employed robust selectors and partial auto-refresh to reset states.
3.  **Multi-OS Support**

    -   *Challenge*: Scripting differences on Mac vs. Windows (permissions, shells, etc.).
    -   *Solution*: Maintained dual script flows, tested on both operating systems thoroughly.
4.  **User Interference**

    -   *Challenge*: If the user scrolled or clicked around, DOM states could shift.
    -   *Solution*: Advised minimal interaction; added safeguards to re-locate elements.

* * * * *

Outcome & Impact
----------------

-   **~400 Leads Captured**: The client reported about 400 successful leads during active use.
-   **1-Second Response Advantage**: Delivered an edge over competitors in a high-stakes sales environment.
-   **Time Saved**: Freed the salesperson from constantly refreshing email and racing to click links.
-   **Cross-Platform Deployment**: Validated on both Mac and Windows, meeting the client's diverse needs.

* * * * *

Lessons Learned
---------------

-   **Fragility of Web Automation**: Reliance on specific UI elements requires close monitoring of any third-party layout changes.
-   **Balancing Speed & Stability**: Cutting down detection latency to ~1 second introduced complexities in maintaining reliability.
-   **Platform Nuances**: Even with a "universal" Node.js solution, handling Mac's and Windows' different shell scripts demanded careful packaging and testing.
