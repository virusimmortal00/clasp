"use strict";
//test
function createCustomerPresentations(customerNames, templateId) {
    for (const customerName of customerNames) {
        const templateFile = DriveApp.getFileById(templateId);
        const newFile = templateFile.makeCopy(`${templateFile.getName()} - ${customerName}`);
        const newPreso = SlidesApp.openById(newFile.getId());
        // Get all slides in the new presentation
        const slides = newPreso.getSlides();
        // Loop over each slide
        for (const slide of slides) {
            // Get all text boxes in the slide
            const shapes = slide.getShapes();
            // Loop over each text box
            for (const shape of shapes) {
                const textRange = shape.getText();
                // Replace {{customer_name}} with the actual customer name
                if (textRange.asString().includes('{{customer_name}}')) {
                    textRange.replaceAllText('{{customer_name}}', customerName);
                }
            }
        }
    }
}
// Call the function
const customerNames = ["Company1", "AsherCompany", "Bobtest"];
const templateId = '1ex5yiHNcWt34y9OyjzpcXCQAjiQtB9sevDQNeVALmjE';
createCustomerPresentations(customerNames, templateId);
console.log("test");
