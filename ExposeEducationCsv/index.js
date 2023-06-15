const puppeteer = require("puppeteer");
const fetch = require('node-fetch');

// Hardwired to start 2022-05-01 to support loading to tables [2022-06-09]

module.exports = async function (context, req) {
    context.log('JavaScript HTTP trigger function processed a request.');
    context.log("User agent: " + req.headers['user-agent']);


    var account = context.bindingData.account;
    context.log("Account:");
    context.log(account);

    user_reference = account + "user";
    pass_reference = account + "password";
    user = process.env[user_reference];
    password = process.env[pass_reference];


    context.log("User");
    context.log(user.slice(0, 5));
    context.log("Password");
    context.log(password.slice(0, 5));

    const urleducation = req.query.url || "https://portal.azure.com/#blade/Microsoft_Azure_Education/EducationMenuBlade/classrooms";
    const browser = await puppeteer.launch({headless: true});
    const page = await browser.newPage();
    await page.goto(urleducation, {waitUntil: 'networkidle2'});
    await page.waitFor(1111);
    
    // Intercept bearer token
    context.log('Begin Intercept bearer token');
    var bearertoken;
    page.on('request', interceptedRequest => {
        let headers = interceptedRequest.headers();
        console.log('\n\n****************************************************\n\n');
        console.log(interceptedRequest.postData());
        const regexBearer = RegExp('Bearer');
        if ( regexBearer.test(headers['authorization']) ) {
            bearertoken = headers['authorization'];
        }
    });

    // User email
    context.log('Begin User email');
    await page.type('#i0116.form-control.ltr_override.input.ext-input.text-box.ext-text-box', user);
    page.keyboard.press(String.fromCharCode(13));
    await page.waitFor(5000);
    await page.waitFor(1111);

    // Password
    context.log('Begin Password');
    await page.type('#i0118.form-control.input.ext-input.text-box.ext-text-box', password);
    page.keyboard.press(String.fromCharCode(13));
    await page.waitFor(5000);
    await page.waitFor(1111);

    // Remember me?
    context.log('Begin Remember me?');
    await page.click('#idSIButton9.button.ext-button.primary.ext-primary');
    await page.waitFor(5000);
    page.keyboard.press(String.fromCharCode(13));
    await page.waitFor(10000);
    page.waitForNavigation({ waitUntil: 'networkidle0' });

    await page.waitFor(7000);

    // Bogus load of edu usage endpoint
    context.log('Begin Bogus load of edu usage endpoint');
    await page.goto('https://frontdoor.educationhub.microsoft.com/api/Usage/GetUsageReport')
    await page.waitFor(4000);
    context.log(bearertoken);


    // csv POST request
    context.log('Begin csv POST request');
    var csvData;

    async function retrieveCsv() {
        csvData = await page.evaluate(({bearertoken}) =>
        {
            return fetch('https://frontdoor.educationhub.microsoft.com/api/Usage/GetUsageReport', {
                method: 'POST',
                headers: { "Authorization": bearertoken, "Content-Type": "application/json", "Accept": "application/json" },
                body: JSON.stringify({"startDateTime":"2023-06-09T00:00:00.007Z","endDateTime":"2026-06-30T00:00:00.007Z"})
            }).then(r => r.text());
        },{bearertoken});
    }
    function checkCsv() {
        let csvDataSample = csvData.slice(0, 100);
        context.log("Sample text from CSV:");
        context.log(csvDataSample);
        //put CSV check here
        return "CSV sampled"
    }

    retrieveCsv();
    await page.waitFor(40000);
    context.log(checkCsv());
    

    var resultFull = csvData;
    var resultFullSample = resultFull.slice(0, 100);
    context.log("Sample text of result:");
    context.log(resultFullSample);

    const regexCourseName = RegExp('CourseName');
    if ( regexCourseName.test(resultFullSample) ) {
        context.log("First result includes expected header");
    } else {
        context.log("First result did not include expected header; trying again");
        retrieveCsv();
        await page.waitFor(3000);
        resultFull = csvData;
        resultFullSample = resultFull.slice(0, 100);
        context.log("Sample text of second try:");
        context.log(resultFullSample);   
    }

    context.res = {
        body: resultFull,
        headers: {
            "content-type": "text/csv"
        }
    };

    await browser.close();
    context.log('Finished.');
    context.done();

};
