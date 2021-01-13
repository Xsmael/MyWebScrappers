const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('main');
workbook.creator = 'TheWebScraper';
workbook.lastModifiedBy = 'TheWebScraper';
worksheet.columns = [
    { header: 'DENOMINATION', key: 'DENOMINATION' },
    { header: 'SIGLE', key: 'SIGLE'},
    { header: 'TEL', key: 'TEL' },
    { header: 'E-mail', key: 'mail' },
    { header: 'GERANT', key: 'GERANT' },
    { header: 'ADRESSE', key: 'ADRESSE' },
    { header: 'CAPITAL', key: 'CAPITAL' },
    { header: 'SIEGE', key: 'SIEGE' },
    { header: 'FORME', key: 'FORME' },
    { header: 'SECTEUR', key: 'SECTEUR' },
    { header: 'OBJET', key: 'OBJET' },
    { header: 'RCCM', key: 'RCCM' },
    { header: 'CREATION', key: 'CREATION' }
  ];

function getDigits(str) { return (str.match(/([0-9]+)/g) || [0] ) [0] ; }

var today= new Date(Date.now());
var strDate= today.getDate()+"-"+(today.getMonth()+1)+"-"+today.getFullYear();
var fileName="bd-"+ strDate +".xlsx";


var startTime= Date.now(); 
(async () => {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    const userAgent = 'Mozilla/5.0 (X11; Linux x86_64)' +
                        'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.39 Safari/537.36';
                        await page.setUserAgent(userAgent);
    var pageId=0;
    do{
        console.log("Loading next page hehe... :P");
        await page.goto('http://www.me.bf/en/annonces-legales?order=field_al_date_rccm&sort=desc&page='+pageId);
        // await page.pdf({path: i+'.pdf', format: 'A4'}); // un comment this to generate PDF files as well
        
        var table = await page.evaluate(() => {  // table(array) represents the table list on a page (www.me.bf/en/annonces-legales) it has 20 rows or items
            var arr=[];
            for(let i= 0; i < 20; i++){
                if( document.querySelectorAll('td.views-field .center.print-this')[i]){
                    arr.push(
                         document.querySelectorAll('td.views-field .center.print-this')[i]
                        .textContent
                        .replace(/[ \f\r\t\v]+/g,' ').trim()
                        .replace(/Siège/g,'Siege').trim() // Because this is part of params list
                        );
                    }
                    else break;
                }
                return arr;
            });

            
            table.forEach(tableItem => {
                var lines= tableItem.split('\n');;
                var params= ['DENOMINATION', 'SIGLE', 'TEL', 'E-mail', 'GERANT','ADRESSE', 'CAPITAL', 'SIEGE', 'FORME', 'OBJET', 'FORMALITES AU GREFFE'];
                var dataObj={};
                for(let line of lines) {
                    for( let param of params) {
                    if(line.trim().toLowerCase().startsWith(param.toLowerCase()) ) {
                        dataObj[param]= dataObj[param] || (line.split(':')[1] || "").trim();

                        if(param == 'OBJET') {
                            dataObj['SECTEUR']= dataObj['SECTEUR'] || (line.split(':')[2] || "").trim();
                        }
                        if(param == 'CAPITAL') {
                            dataObj[param] = getDigits(dataObj[param]);
                        }
                        if(param=='FORMALITES AU GREFFE') {
                            let dateRE= /(\d{2}\/\d{2}\/\d{4})/;
                            dataObj['CREATION'] = dateRE.exec(dataObj[param])[1];
                        }
                        break;
                    }
                    else if(line.trim().toLowerCase().startsWith("sous le numéro") ) {
                        let rccmRE= /(BF[A-Z]+[0-9A-Z]+)/;
                        dataObj['RCCM'] =     (rccmRE.exec(line.trim()) || ["","NULL"] )[1];         
                    }
                }
            }
            worksheet.addRow(dataObj);
            console.log(dataObj);
        });
        console.log('+++++++++++++++++++++'+'PAGE NO '+pageId+' ++++++++++++++++++++++++++');
        // console.log(textContent);
        // console.log(await page.evaluate('document.querySelectorAll("td.views-field .center.print-this")[0]')  );
        // console.log(await page.evaluate('document.querySelectorAll("td.views-field .center.print-this .text-left b")[1]')  );
        // await browser.close();
        pageId++;
    } while (table.length >0);
    
    var endTime= Date.now();  
    var timeElapsed= ((endTime-startTime)/1000) /60 ; // in minutes
    worksheet.addRow({
        'DENOMINATION':"Generated on "+ strDate+ " in " +timeElapsed.toFixed(2)+" minutes\n "+pageId+" pages processed"
    });
    worksheet.addRow({
        'DENOMINATION':"THIS MATERIAL IS PROVIDED AS IS, WITH ABSOLUTELY NO WARRANTY EXPRESSED OR IMPLIED.  ANY USE IS AT YOUR OWN RISK."
    });
    await workbook.xlsx.writeFile(fileName);
    console.log("JOB FREAKIN DONE !");
    console.log("started at "+ Date(startTime));
    console.log("ended   at "+ Date(endTime));
    console.log(pageId+" pages parsed in "+ timeElapsed.toFixed(2) +" min.");
    console.log("That is "+ (timeElapsed*60/pageId).toFixed(2) + " seconds per page");
    
})();
