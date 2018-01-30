var less = require('less');
var fs = require('fs');
//compile and add stylesheets
var css = document.createElement('style');
css.type = 'text/css';
less.render('@import "stylesheets/style.less"; ',
    function (e, output) {
        css.textContent = output.css;
        document.getElementsByTagName('head')[0].appendChild(css);
    });
var XLSX = require('xlsx');
var dragNDrop;
var uploadHTML;
var sheetOption = "All";
var source = "";
var destination = "";
var fileName= "";
var processing = false;
window.onload = function () {
    dragNDrop = document.getElementById('drag');
    uploadHTML = dragNDrop.innerHTML;
    dragNDrop.ondragover = () => {
        dragNDrop.classList.add('fileHover');
        return false;
    }
    dragNDrop.ondragleave = function () {
        dragNDrop.classList.remove('fileHover');
        return false;
    };

    dragNDrop.ondragend = function () {
        return false;
    };

    dragNDrop.ondrop = (e) => {
        e.preventDefault();
        var flag = true;
        if (e.dataTransfer.files.length > 1) dragNDrop.innerHTML = uploadHTML + "<p>Max 1 file allowed.</p>"
        else {
            flag = verifyFile(e.dataTransfer.files[0].path);
            if (flag) {
                dragNDrop.innerHTML = uploadHTML + "<p>" + e.dataTransfer.files.length + " file selected.</p>";
                source = e.dataTransfer.files[0].path;
                destination = input.files[0].path.substring(0, input.files[0].path.lastIndexOf('/'))
                document.getElementById("destination").nextElementSibling.innerHTML = "..." + destination.substring(destination.lastIndexOf('/')) +" | <b>OR Click to Change</b>"
            }
            else{
                 dragNDrop.innerHTML = uploadHTML + "<p>Only .xlsx is allowed.</p>";
                source= "";
            }
        }
        dragNDrop.classList.remove('fileHover');
        return false;
    };
}

function chooseFiles(input) {
    var flag = verifyFile(input.files[0].path)
    if (flag) {
        dragNDrop.innerHTML = uploadHTML + "<p>" + input.files.length + " file's selected. Name: "+input.files[0].name+"</p>"
        source = input.files[0].path;
        destination = input.files[0].path.substring(0, input.files[0].path.lastIndexOf('/'))
        document.getElementById("destination").nextElementSibling.innerHTML = "..." + destination.substring(destination.lastIndexOf('/')) +" | <b>OR Click to Change</b>"
    } 
    else{ 
        dragNDrop.innerHTML = uploadHTML + "<p>Only .xlsx is allowed.</p>"
        source = "";
    }
}

function changeSheets(value) {
    document.getElementById("lowerRange").disabled = true;
    document.getElementById("upperRange").disabled = true;
    sheetOption = value;
}

function enableRange(value) {
    document.getElementById("lowerRange").disabled = false;
    document.getElementById("upperRange").disabled = false;
    sheetOption = value;
}

function toggleSheets(value) {
    if (sheetOption != value) {
        var curr = document.getElementById("lowerRange").disabled;
        document.getElementById("lowerRange").disabled = !curr;
        document.getElementById("upperRange").disabled = !curr;
        sheetOption = value;
    }
}

function chooseDestination(input) {
    destination = input.files[0].path;
    input.nextElementSibling.innerHTML = "....." + destination.substring(destination.lastIndexOf('/'));
}

function verifyFile(path) {
    return path.endsWith('.xlsx');
}

function loading(){
    console.log(document.getElementsByClassName('loader')[0])
    if (processing){
        document.getElementsByClassName('loader')[0].classList.remove('hidden');
        setTimeout(loading, 1000);
    }
    else{
        document.getElementsByClassName('loader')[0].classList.add('hidden');
        document.getElementById("message").classList.remove('hidden');
    }
}

function generate() {
    var separator = document.getElementById("separator");
    var sheetRange = [document.getElementById("lowerRange").value, document.getElementById("upperRange").value]
    if (source == "") document.getElementById("message").innerHTML = "<p>Error : No file selected.</p>";
    else if (sheetOption == "Range" && (sheetRange[0]< 0 || sheetRange[1] < 0 || sheetRange[0] == "" ||
        sheetRange[1] == "")) document.getElementById("message").innerHTML = "<p>Error : Invalid Sheet Range.</p>";
    else{
        try{
            processing = true;
            loading();
            parseFile(separator.value, [Math.min(sheetRange[0],sheetRange[1]),Math.max(sheetRange[0],sheetRange[1])]);
        }
        catch(e){
            console.log('An error has occurred: '+e.message)
            processing = false;
            document.getElementById("message").classList.add('hidden');
            document.getElementById("message").innerHTML = "<p>Error: Try a different separator</p>";
            //convert to string and use replaceAll to overcome limited separator barriers

        }
    }
}

function parseFile(separator, sheetRange) {
    var workbook = XLSX.readFile(source);
    var parseString = "";
    if(sheetOption == "All"){
        workbook.SheetNames.forEach(wb => {
            parseString += "Sheet : "+wb+"\n" + XLSX.utils.sheet_to_csv(workbook.Sheets[wb],{FS:separator})+"\n";
        })
    }
    else {
        for(i=sheetRange[0]; i<= sheetRange[1];i++){
            wb = workbook.SheetNames[i-1];
            if(wb != undefined) parseString += "Sheet : "+wb+"\n" + XLSX.utils.sheet_to_csv(workbook.Sheets[wb],{FS:separator})+"\n";
            else break;
        }
    }
    saveFile(parseString)
    console.log(parseString, sheetRange)
}

function saveFile(parseString){
    fs.writeFile(destination+"/"+source.substring(source.lastIndexOf('/'), source.lastIndexOf('.'))+".txt" , parseString, (err) => {
        if(err){
            console.log("An error ocurred creating the file "+ err.message)
        }
        processing = false;
        document.getElementById("message").classList.add('hidden');
        document.getElementById("message").innerHTML = "<p>Succesfull! File saved at "+destination+" </p>";    
    });
}