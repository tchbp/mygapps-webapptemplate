let idnwsplogo = "1xUdhjmLjVwD742QImsa3snzXUDA3J6ny"; //Id Logo .jpg, .png
let wsData = "1z37VXmy21xy0gIbGlIpHslzrLsctn8pxN9jD6ozP-0I";//Id Sheet Data

function doGet() {
    let htmlout = HtmlService.createTemplateFromFile('index');
    htmlout.logo = DriveApp.getFileById(idnwsplogo).getDownloadUrl();
//เพิ่ม Code
    return htmlout.evaluate()
                .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
