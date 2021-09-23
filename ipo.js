//npm i puppeteer -> Install puppeteer Module
//npm i path -> Install Path Module
//npm i require
//npm i promise
//npm i xlsx

// Enter you Login credentials in login.js code to run this code.

//URL -> "https://kite.zerodha.com/" 

let puppeteer = require("puppeteer");
let path = require("path");
let fs = require("fs");
let LoginDetails = require("./login");
let xlsx = require("xlsx");
const { Http2ServerRequest } = require("http2");

let page, browser;

(async function IpoInfo(){

    try{

        let browserOpenPromise = puppeteer.launch({
            headless: null,
            // slowMo: 2000,
            defaultViewport: null,
            args:["--start-maximized", "--disable-notification"]
        })

        let browserObj = await browserOpenPromise;
        console.log("Browser Opened");

        let browser = browserObj;

        page = await browserObj.newPage();
        await page.goto("https://kite.zerodha.com/");
        await page.waitFor(2000);
        await page.type("[placeholder='User ID (eg: AB0001)']", LoginDetails.USERID, {delay: 100});
        await page.type("[placeholder='Password']", LoginDetails.PASSWORD, {delay: 100});
        await page.keyboard.press("Enter");
        await page.waitFor(2000);
        await page.type("#pin", LoginDetails.PIN, {delay: 100});
        await page.keyboard.press("Enter");
        await WaitAndClick("[class='user-id']", page)
        await WaitAndClick("[href='https://console.zerodha.com/dashboard/']", page);
        await page.waitFor(2000);
        
        let TArray = await browser.pages();
        console.log("NoofTabs: " + TArray.length);
        let RTab = TArray[TArray.length - 1];

        await WaitAndClick("[class='portfolio-id']", RTab, {delay: 100});
        await WaitAndClick("[href='/portfolio/ipo']", RTab, {delay: 100});
        await page.waitFor(2000);

        // Geting Heading Array Of IPOs
        await RTab.waitForSelector("h2", {visible: true});
        let StatusOfIPO = await RTab.$$("h2");

        // -----------------------------------------------------------------------------------------------------------------------------------------
        
        // 1st Element of IPOs
        let OngoingIPOs = await RTab.evaluate(el => el.textContent, StatusOfIPO[0]);
        console.log(OngoingIPOs.trim());

        // Ongoing IPOs Table
        let ArrayOfIPOs = await RTab.$$("tbody");

        console.log("No.of OnGoingIPOs: ", ArrayOfIPOs.length);

        for( let i = 1; i <= ArrayOfIPOs.length; i++){

            let IPODetailsArr = await RTab.$$("td");

            let Company_Name = await RTab.evaluate(el => el.textContent, IPODetailsArr[i * 8 - 1]);
            let Start_Date = await RTab.evaluate(el => el.textContent, IPODetailsArr[i * 8 ]);
            let End_Date = await RTab.evaluate(el => el.textContent, IPODetailsArr[i * 8 + 1]);
            let Price_Range = await RTab.evaluate(el => el.textContent, IPODetailsArr[i * 8 + 2]);
            let Minimum_Quantitiy = await RTab.evaluate(el => el.textContent, IPODetailsArr[i * 8 + 3]);
            let Status = await RTab.evaluate(el => el.textContent, IPODetailsArr[i * 8 + 4]);

            ProsessIPOs(OngoingIPOs.trim(), Company_Name.trim(), Start_Date, End_Date, Price_Range.trim(), Minimum_Quantitiy, Status ); 

        }

        // Opening more deatils of Upcoming IPOs and Closed IPOs.

        await WaitAndClick("[href='https://zerodha.com/ipo']", RTab, {delay: 100});
        await page.waitFor(2000);

        let T2Array = await browser.pages();
        console.log("No.of Tabs: " + T2Array.length);
        let R2Tab = T2Array[T2Array.length - 1]

        // Geting Heading Array Of IPOs
        await R2Tab.waitForSelector("h2", {visible: true});
        let StatusOfIPO2 = await R2Tab.$$("h2");

        // -----------------------------------------------------------------------------------------------------------------------------------------

        // 2nd Element Of IPO
        let UpcomingIPOs = await R2Tab.evaluate(el => el.textContent, StatusOfIPO2[1]);
        console.log(UpcomingIPOs.trim());

        // Upcoming IPO's Table
        let UpcomingIPOArr = await R2Tab.$$("#ipo > div:nth-child(2) > table > tbody > tr");
        console.log("No.of Upcoming IPOs: " + UpcomingIPOArr.length);

        for( let i = 1; i <= UpcomingIPOArr.length; i++ ){

            let IPODetaislArr2 = await R2Tab.$$("#ipo > div:nth-child(2) > table > tbody > tr > td");

            let Company_Name = await R2Tab.evaluate(el => el.textContent, IPODetaislArr2[i * 5 - 5]);
            let Date_Open_Close = await R2Tab.evaluate(el => el.textContent, IPODetaislArr2[i * 5 - 4]);
            let Price_Range = await R2Tab.evaluate(el => el.textContent, IPODetaislArr2[i * 5 - 3]);
            let Minimum_Quantitiy = await R2Tab.evaluate(el => el.textContent, IPODetaislArr2[i * 5 - 2]);

            ProsessIPOs2(UpcomingIPOs.trim(), Company_Name, Date_Open_Close, Price_Range, Minimum_Quantitiy);

        }

        // -----------------------------------------------------------------------------------------------------------------------------------------

        // 3rd Element of IPO
        let ClosedIPOs = await R2Tab.evaluate(el => el.textContent, StatusOfIPO2[2]);
        console.log(ClosedIPOs.trim());

        // Closed IPO's Table
        let ClosedIPOArr = await R2Tab.$$("#ipo > div:nth-child(4) > table > tbody > tr");
        console.log("No.of Upcoming IPOs: " + ClosedIPOArr.length);

        for( let i = 1; i <= ClosedIPOArr.length; i++ ){

            let IPODetaislArr3 = await R2Tab.$$("#ipo > div:nth-child(4) > table > tbody > tr > td");

            let Company_Name = await R2Tab.evaluate(el => el.textContent, IPODetaislArr3[i * 5 - 5]);
            let Date_Open_Close = await R2Tab.evaluate(el => el.textContent, IPODetaislArr3[i * 5 - 4]);
            let Price_Range = await R2Tab.evaluate(el => el.textContent, IPODetaislArr3[i * 5 - 3]);
            let Minimum_Quantitiy = await R2Tab.evaluate(el => el.textContent, IPODetaislArr3[i * 5 - 2]);

            ProsessIPOs2(ClosedIPOs.trim(), Company_Name.trim(), Date_Open_Close, Price_Range, Minimum_Quantitiy);

        }
        

    }catch (err){
        console.log(err);
    }

})();

let FolderPath = process.cwd();
let IPOFolderPath = path.join(FolderPath, "IPOFolder");
fs.mkdirSync(IPOFolderPath);

function ProsessIPOs(StatusNameOfIPO, Company_Name, Start_Date, End_Date, Price_Range, Minimum_Quantitiy, Status){
    
    let Obj = {
        Company_Name,
        Start_Date,
        End_Date,
        Price_Range,
        Minimum_Quantitiy,
        Status
    }

    let IPOFilePath = path.join(IPOFolderPath, StatusNameOfIPO + ".xlsx");
    let IPODetails = [];


    if(fs.existsSync(IPOFilePath) == false){
        IPODetails.push(Obj);
    }else{
        IPODetails = excelReader(IPOFilePath, StatusNameOfIPO)
        IPODetails.push(Obj);
    }

    excelWriter(IPOFilePath, IPODetails, StatusNameOfIPO );


}

function ProsessIPOs2(StatusNameOfIPO, Company_Name, Date_Open_Close, Price_Range, Minimum_Quantitiy){
    
    let Obj = {
        Company_Name,
        Date_Open_Close,
        Price_Range,
        Minimum_Quantitiy,
    }

    let IPOFilePath = path.join(IPOFolderPath, StatusNameOfIPO + ".xlsx");
    let IPODetails = [];


    if(fs.existsSync(IPOFilePath) == false){
        IPODetails.push(Obj);
    }else{
        IPODetails = excelReader(IPOFilePath, StatusNameOfIPO)
        IPODetails.push(Obj);
    }

    excelWriter(IPOFilePath, IPODetails, StatusNameOfIPO );


}

function excelReader(filePath, sheetName) {
    // workbook
    let wb = xlsx.readFile(filePath);
    // get data from a particular sheet in that wb
    let excelData = wb.Sheets[sheetName];
    // sheet to json 
    let ans = xlsx.utils.sheet_to_json(excelData);
    return ans;
}

function excelWriter(filePath, json, sheetName) {
    // workbook create
    let newWB = xlsx.utils.book_new();
    // worksheet
    let newWS = xlsx.utils.json_to_sheet(json);
    xlsx.utils.book_append_sheet(newWB, newWS, sheetName);
    // excel file create 
    xlsx.writeFile(newWB, filePath);
}

function WaitAndClick(selector, Spage){

    return new Promise(function(resolve, reject){
        let WaitfortheselectorPromise = Spage.waitForSelector(selector, {visilbe: true});

        WaitfortheselectorPromise.then(function(){

            let ClicktheSelectorPromise = Spage.click(selector, {delay: 100});
            return ClicktheSelectorPromise;

        }).then(function(){
            resolve();
        }).catch(function(err){
            reject(err);
        })
    })
}

