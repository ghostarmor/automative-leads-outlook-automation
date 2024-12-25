
const puppeteer = require('puppeteer');

const processedEmails = new Set();

const leadSenderEmail = 'noreply@leadmanager.com'
const respondToEmailHref= 'https://example-href.com'



async function monitorOutlook() {
    const browser = await puppeteer.launch({ headless: false, args: ["--disable-notifications"] });
    const context = browser.defaultBrowserContext();
    context.overridePermissions("https://outlook.live.com/mail/0/", ["notifications"]);
    const page = await browser.newPage();
    await page.goto('https://outlook.office.com/');

    // Wait for the user to log in manually if needed
    await page.waitForNavigation({ waitUntil: 'networkidle2' });
    page.on('dialog', async dialog => {
        await dialog.accept();
    });

    
    // Wait for the inbox to loadawait page.waitForSelector('div.customScrollBar.jEpCF', { timeout: 0 });

    // Expose handleNewEmail to Node.js
    await page.exposeFunction('handleNewEmail', async (nodeHtml, emailId) => {
        if (!processedEmails.has(emailId)) {
            processedEmails.add(emailId);
            await handleNewEmail(nodeHtml);
        }
    });

    // Observe the inbox for new emails
    await page.evaluate(() => {
        var sessionStart = new Date(new Date().toLocaleString("en-US", { timeZone: "America/Denver" }));

        const targetNode = document.querySelector('div.customScrollBar.jEpCF > div');
        if (!targetNode) {
            console.log('Inbox not found.');
            return;
        }

        function parseTimestampToMinutes(timestamp) {
            const [time, period] = timestamp.split(' ');
            let [hour, minute] = time.split(':').map(Number);

            // Convert to 24-hour format
            if (period === 'PM' && hour !== 12) {
                hour += 12;
            } else if (period === 'AM' && hour === 12) {
                hour = 0;
            }

            // Convert the time to total minutes since midnight
            return hour * 60 + minute;
        }

        function isTimestampAfterDate(timestamp, date) {
            const timestampMinutes = parseTimestampToMinutes(timestamp);

            // Convert the specific date's time to total minutes since midnight
            const specificDateMinutes = date.getHours() * 60 + date.getMinutes();

            console.log(timestampMinutes);
            console.log(specificDateMinutes);

            return timestampMinutes >= specificDateMinutes;
        }

        function waitForElement(parent, selector, callback) {
            if (parent.querySelector(selector)) {
                callback();
            } else {
                console.log('Trying to fetch...');
                setTimeout(() => waitForElement(parent, selector, callback), 500);
            }
        }

        const config = { childList: true, subtree: true };

        const callback = function (mutationsList, observer) {
            for (let mutation of mutationsList) {
                if (mutation.type === 'childList') {
                    mutation.addedNodes.forEach(node => {
                        if (node.nodeType === 1 && node.querySelector && node.querySelector(`span[title=${leadSenderEmail}]`)) {
                            waitForElement(node, "span._rWRU.Ejrkd.qq2gS", () => {
                                const timestamp = node.querySelector("span._rWRU.Ejrkd.qq2gS");
                                console.log(timestamp.innerText);
                                const dayPattern = /^[A-Za-z]{3} \d{1,2}:\d{2} (AM|PM)$/;
                                if (!dayPattern.test(timestamp.innerText)) {
                                    console.log('From today!');
                                    //Timestamp from today
                                    if (isTimestampAfterDate(timestamp.innerText, sessionStart)) {
                                        console.log(node);
                                        node.firstChild.firstChild.click();
                                        waitForElement(document, ".Q8TCC.yyYQP.customScrollBar", () => {
                                            const emailNode = document.querySelector('.Q8TCC.yyYQP.customScrollBar');
                                            if (emailNode) {
                                                const link = emailNode.querySelector(`a[href^=${respondToEmailHref}]`);
                                                link.click();
                                            } else {
                                                console.log('Email node not found.');
                                            }
                                        });
                                    }
                                }
                            });

                        }
                    });
                }
            }
        };

        const observer = new MutationObserver(callback);
        observer.observe(targetNode, config);
        console.log('MutationObserver setup complete.');
    });

    console.log('Monitoring for new emails.');
}

async function handleNewEmail(emailNodeHtml) {
    console.log('Handling new email...');
    const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();
    await page.setContent(emailNodeHtml);
    await page.screenshot({ path: 'email_screenshot.png' });

    setTimeout(async () => {
        await browser.close();
    }, 10000); 
}

monitorOutlook().catch(error => {
    console.error('Error in monitorOutlook:', error);
});

// shutdown on SIGINT (Ctrl+C)
process.on('SIGINT', async () => {
    console.log('Received SIGINT. Shutting down...');
    if (browser) {
        await browser.close();
    }
    process.exit(0);
});