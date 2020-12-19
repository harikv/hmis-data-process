import * as puppeteer from 'puppeteer';
import * as fs from 'fs';
import * as _ from 'lodash';

const MONTH_ARRAY = ['April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December', 'January', 'February', 'March'];

interface Status {
 [index: string]: {
     [index: string]: boolean
 }
}

const timeout = (ms: number): Promise<void> => {
    return new Promise(resolve => setTimeout(resolve, ms));
}

const initBrowser = async (): Promise<puppeteer.Browser> => {
    console.log('Setting up Browser...');
    const options: puppeteer.LaunchOptions = {
        headless: true,
        timeout: 0,
        defaultViewport: null,
        ignoreHTTPSErrors: true         
    };
    const browser: puppeteer.Browser = await puppeteer.launch(options);
    return browser;
};

const initPage = async (browser: puppeteer.Browser): Promise<puppeteer.Page> => {
    const page = await browser.newPage();
    page.on('console', (consoleMessageObject) => {
        if (consoleMessageObject.type() !== 'warning') {
            console.debug(consoleMessageObject.text())
        }
    });
    return page;
}

const getDataFiles = async (browser: puppeteer.Browser, year: string, state: string): Promise<boolean> => {  
    console.log(`Downloading data for FY${year} for state ${state}`);  
    const page = await initPage(browser);
    const year_regex_match = year.match(/(\d+)-(\d+)/);
    let year1 = '';
    let year2 = '';
    if (year_regex_match.length == 3) {
        year1 = year_regex_match[1];
        year2 = year_regex_match[2];
    }

    try {
        // Visit directory page;
        await page.goto('http://nrhm-mis.nic.in/hmisreports/frmstandard_reports.aspx');
        // Visit sub directory
        await page.waitForSelector('input#ctl00_ContentPlaceHolder1_gridDirList_ctl06_imgDir');
        await page.click('input#ctl00_ContentPlaceHolder1_gridDirList_ctl06_imgDir');
        await page.waitForSelector('input#ctl00_ContentPlaceHolder1_gridDirList_ctl03_imgDir');
        await page.click('input#ctl00_ContentPlaceHolder1_gridDirList_ctl03_imgDir');
        // Select appropriate year
        await page.waitForSelector('table#ctl00_ContentPlaceHolder1_gridDirList');
        // Get ID of appropriate Year
        const elemID = await page.$$eval('table#ctl00_ContentPlaceHolder1_gridDirList > tbody > tr', (rows, year) => {
            const selectedRow = rows.find((row) => { 
                const cell = row.querySelector('td:nth-child(2)');
                if (cell !== null) {
                    const content = cell.textContent.trim();
                    return content == year;
                } else {
                    return false;
                }
            });
            if (selectedRow !== undefined) {
                return selectedRow.querySelector('td:nth-child(1) > input').id;
            } else {
                return undefined;
            }
        }, year);        
        await page.click(`input#${elemID}`);
        // Go to monthwise state directory
        await page.waitForSelector('input#ctl00_ContentPlaceHolder1_gridDirList_ctl02_imgDir');
        await page.click('input#ctl00_ContentPlaceHolder1_gridDirList_ctl02_imgDir');
        // Go to appropriate state
        await page.waitForSelector('table#ctl00_ContentPlaceHolder1_gridDirList');        
        // Get ID of appropriate Year
        const stateID = await page.$$eval('table#ctl00_ContentPlaceHolder1_gridDirList > tbody > tr', (rows, state) => {
            const selectedRow = rows.find((row) => { 
                const cell = row.querySelector('td:nth-child(2)');
                if (cell !== null) {
                    const content = cell.textContent.trim();
                    return content == state;
                } else {
                    return false;
                }
            });
            if (selectedRow !== undefined) {
                return selectedRow.querySelector('td:nth-child(1) > input').id;
            } else {
                return undefined;
            }
        }, state);        
        await page.click(`input#${stateID}`);
        // Get all download links from first page
        await page.waitForSelector('table#ctl00_ContentPlaceHolder1_gridFileList');
        const [downloadLinks, downloadNames] = await page.$$eval('table#ctl00_ContentPlaceHolder1_gridFileList > tbody > tr', (rows) => {            
            const selectedRows = rows.filter((row) => { 
                const cell = row.querySelector('td:nth-child(1) > input');
                return cell !== null;
            });                        
            // Return both links and file names as a zipped array 
            return [selectedRows.map(row => { return row.querySelector('td:nth-child(1) > input').id}), selectedRows.map(row => { return row.querySelector('td:nth-child(2)').textContent.trim()})];
        });
        const downloadList = _.zip(downloadLinks, downloadNames);
        if (downloadList.length > 0) {
            // Go through all file links
            for (const [link, fileName] of downloadList) {
                await page.$eval(`input#${link}`, elem => { 
                    const htmlelem = elem as HTMLElement;
                    htmlelem.click();
                });                
                await page.waitForTimeout(1000);
                // Go to open page
                const tabs = await browser.pages();
                const numTabs = tabs.length;                
                const downloadPage = tabs[numTabs - 1];
                await downloadPage.bringToFront();
                // Set download behaviour
                const client = await downloadPage.target().createCDPSession();
                await client.send('Page.setDownloadBehavior', {behavior: 'allow', downloadPath: './data'});
                // Click download link
                await downloadPage.waitForSelector('input#lbFile');
                await downloadPage.$eval(`input#lbFile`, elem => { 
                    const htmlelem = elem as HTMLElement;
                    htmlelem.click();
                });
                // Check if filename exists in download directory
                let fileExists = false;
                let attempts = 1;
                do {                
                    const files = await fs.promises.readdir(`./data`);                
                    fileExists = files.includes(fileName);
                    attempts += 1;
                    await timeout(500);
                } while (fileExists !== true);                
                await downloadPage.close();                
                // Get month
                const matches = fileName.match(/[a-zA-Z0-9&\s-]+_([a-zA-Z]+)\.xls/);
                let month = '';
                if (matches.length > 0) {
                    month = matches[1];
                }
                // Set to appropriate year
                const monthIndex = MONTH_ARRAY.indexOf(month);
                const yearName = monthIndex > 8 ? year2 : year1;                
                // Rename file to {Year_state_month}
                await fs.promises.rename(`./data/${fileName}`, `./data/${yearName}_${state}_${month}.xls`);
                console.log(`Downloaded data for FY${year} for state ${state} for month ${month}`);
            }
        }
        // Check if there are more results              
        const nextPage = await page.$('table#ctl00_ContentPlaceHolder1_gridFileList tr > td > table tr > td > a');
        if (nextPage !== null) {
            await page.click('table#ctl00_ContentPlaceHolder1_gridFileList tr > td > table tr > td > a');
            await page.waitForSelector('table#ctl00_ContentPlaceHolder1_gridFileList');
            const [downloadLinks, downloadNames] = await page.$$eval('table#ctl00_ContentPlaceHolder1_gridFileList > tbody > tr', (rows) => {                
                const selectedRows = rows.filter((row) => { 
                    const cell = row.querySelector('td:nth-child(1) > input');
                    return cell !== null;
                });                            
                // Return both links and file names as a zipped array 
                return [selectedRows.map(row => { return row.querySelector('td:nth-child(1) > input').id}), selectedRows.map(row => { return row.querySelector('td:nth-child(2)').textContent.trim()})];
            });
            const downloadList = _.zip(downloadLinks, downloadNames);
            if (downloadList.length > 0) {
                // Go through all file links
                for (const [link, fileName] of downloadList) {
                    await page.$eval(`input#${link}`, elem => { 
                        const htmlelem = elem as HTMLElement;
                        htmlelem.click();
                    });                
                    await page.waitForTimeout(1000);
                    // Go to open page
                    const tabs = await browser.pages();
                    const numTabs = tabs.length;                    
                    const downloadPage = tabs[numTabs - 1];
                    await downloadPage.bringToFront();
                    // Set download behaviour
                    const client = await downloadPage.target().createCDPSession();
                    await client.send('Page.setDownloadBehavior', {behavior: 'allow', downloadPath: './data'});
                    // Click download link
                    await downloadPage.waitForSelector('input#lbFile');
                    await downloadPage.$eval(`input#lbFile`, elem => { 
                        const htmlelem = elem as HTMLElement;
                        htmlelem.click();
                    });
                    // Check if filename exists in download directory
                    let fileExists = false;
                    let attempts = 1;
                    do {                
                        const files = await fs.promises.readdir(`./data`);                
                        fileExists = files.includes(fileName);
                        attempts += 1;
                        await timeout(500);
                    } while (fileExists !== true);                
                    await downloadPage.close();                
                    // Get month
                    const matches = fileName.match(/[a-zA-Z0-9&\s-]+_([a-zA-Z]+)\.xls/);
                    let month = '';
                    if (matches.length > 0) {
                        month = matches[1];
                    }
                    // Set to appropriate year
                    const monthIndex = MONTH_ARRAY.indexOf(month);
                    const yearName = monthIndex > 8 ? year2 : year1;                
                    // Rename file to {Year_state_month}
                    await fs.promises.rename(`./data/${fileName}`, `./data/${yearName}_${state}_${month}.xls`);
                    console.log(`Downloaded data for FY${year} for state ${state} for month ${month}`);
                }
            }
        }
        await page.close();
        return true;
    } catch (e) {
        console.error('Error encountered');
        console.error(e);        
        return false;
    }    
}

const main = async () => {    
    const browser = await initBrowser();
    // Set up years
    const yearsToCrawl = ['2017-2018', '2018-2019', '2019-2020', '2020-2021'];
    // const testYears = ['2017-2018'];
    // Set up states
    const states = ["A & N Islands", "Andhra Pradesh", "Arunachal Pradesh", "Assam", "Bihar", "Chandigarh", "Chhattisgarh", "Dadra & Nagar Haveli", "Daman & Diu", "Delhi", "Goa", "Gujarat", "Haryana", "Himachal Pradesh", "Jammu & Kashmir", "Jharkhand", "Karnataka", "Kerala", "Lakshadweep", "Madhya Pradesh", "Maharashtra", "Manipur", "Meghalaya", "Mizoram", "Nagaland", "Odisha", "Puducherry", "Punjab", "Rajasthan", "Sikkim", "Tamil Nadu", "Telangana", "Tripura", "Uttar Pradesh", "Uttarakhand", "West Bengal"];
    // set up tracking for each year and each state
    const status: Status = {};
    // const testStates = ["A & N Islands"];
    for (const year of yearsToCrawl) {
    // for (const year of testYears) {
        for (const state of states) {
        // for (const state of testStates) {
            const success = await getDataFiles(browser, year, state);            
            if (success == true) {
                console.log(`Downloaded all data for FY${year} for state ${state} successfully!`);
            } else {
                if (_.has(status, year)) {
                    status[year][state] = false;
                } else {
                    status[year] = {};
                    status[year][state] = false;
                }
            }
        }
    }
    console.log(`Failures: `);
    console.log(status);
    console.log(JSON.stringify(status));
    await browser.close();    
};

main();