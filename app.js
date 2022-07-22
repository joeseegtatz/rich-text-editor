//Elements
const elements = document.querySelectorAll('.btn');

// Event
elements.forEach(eachElement => {
    eachElement.addEventListener('click', () => {
        let command = eachElement.dataset['element'];
        
        if(command == "export-word"){
            exportHTML();
        }
        else{
            document.execCommand(command, false, null);
        }
        
    });
});

// export HTML
function exportHTML() {
    var header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' " +
        "xmlns:w='urn:schemas-microsoft-com:office:word' "
    var footer = "</body></html>";
    var sourceHTML = header + document.getElementById("source-html").innerHTML + footer;

    var source = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(sourceHTML);
    var fileDownload = document.createElement("a");
    document.body.appendChild(fileDownload);
    fileDownload.href = source;
    fileDownload.download = 'document.doc';
    fileDownload.click();
    document.body.removeChild(fileDownload);
}
