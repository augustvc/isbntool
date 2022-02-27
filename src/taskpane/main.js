let optionArray = ["Leeg", "Auteur", "Volledige Titel", "Uitgever", "Uitgeefdatum", "Huidige datum", "ISBN", "Aantal pagina's", "Korte titel", "Subtitel", "Omschrijving", "Taal"];

let fieldCount = 0;

let createDropdown = function(parent) {
    fieldCount++;
    let txt = document.createElement("p");
    txt.innerHTML = "<b>Veld " + fieldCount + "</b><br>";
    
    parent.appendChild(txt);
    let slct = document.createElement("select");
    
    txt.appendChild(slct);
    for(let i = 0; i < optionArray.length; i++) {
        let option = document.createElement("option");
        option.value = optionArray[i];
        option.text = optionArray[i];
        slct.appendChild(option);
    }
}

let addBtn = document.getElementById("addField");
let removeBtn = document.getElementById("removeField");

let fieldsObj = document.getElementById("fields");

addBtn.addEventListener("click", () => {
    createDropdown(fieldsObj);
});

createDropdown(fieldsObj);
createDropdown(fieldsObj);
createDropdown(fieldsObj);