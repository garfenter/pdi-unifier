const Xlsx = require('xlsx');
const fs = require('fs');

const skillSelected = (sheet, letter) => {
    for (let i = 10; i < 16; i++) {
        let key = `${letter}${i}`;
        let cell = sheet[key] ? sheet[key].v : null;
        if (cell === 'X' || cell === 'x') {
            return true;
        }
    }
    return false;
}

const skills = (sheet, letter) => {
    let result = '';
    for (let i = 10; i < 16; i++) {
        let key = `${letter}${i}`;
        let cell = sheet[key] ? sheet[key].v : null;
        if (cell) {
            result += ', ' + cell;
        }
    }
    return result;
}

const readFile = (file) => {
    let result = {};
    let workbook = Xlsx.readFile(file);
    let sheet = workbook.Sheets['A. PDI INDIVIDUAL'];
    result.file = file;
    result.nombre = sheet['B10'].v;
    result.area = sheet['C10'].v;
    result.fortaleza1 = sheet['F10'].v;
    result.fortaleza2 = sheet['F12'].v;
    result.fortaleza3 = sheet['F14'].v;
    result.enfoque1 = sheet['H10'].v;
    result.enfoque2 = sheet['H12'].v;
    result.enfoque3 = sheet['H14'].v;
    result.plazo1 = sheet['I10'].v;
    result.plazo2 = sheet['J10'].v;
    result.plazo3 = sheet['K10'].v;
    result.otros = sheet['E10'].v;
    result.auditoriaIso = skillSelected(sheet, 'O');
    result.certificacionesVisa = skillSelected(sheet, 'P');
    result.certificacionesTecnologicas = skillSelected(sheet, 'Q');
    result.itil = skillSelected(sheet, 'R');
    result.pci = skillSelected(sheet, 'S');
    result.gestionProyectos = skillSelected(sheet, 'T');
    result.seguridadInformatica = skillSelected(sheet, 'U');
    result.conocimientoNegocio = skillSelected(sheet, 'V');
    result.normasEstandares = skillSelected(sheet, 'W');
    result.sac = skillSelected(sheet, 'X');
    result.forosYComitesExternos = skillSelected(sheet, 'Y');
    result.summit = skillSelected(sheet, 'Z');
    result.ingles = skillSelected(sheet, 'AA');
    result.excel = skillSelected(sheet, 'AB');
    result.otrosCursos = skills(sheet, 'AC');
    return result;
}

const writeResults = (results) => {
    let writter = fs.createWriteStream('out.csv', {
        flags: 'a' // 'a' means appending (old data will be preserved)
    });
    writter.write(`"Archivo", "Nombre", "Area", "Fortaleza 1", "Fortaleza 2", "Fortaleza 3", "Enfoque 1", "Enfoque 2", "Enfoque 3", "Plazo 1", "Plazo 2", "Plazo 3", "Otros", "Auditoria ISO", "Certificaciones VISA", "Certificaciones Tecnologicas", "itil", "PCI", "Gestion de Proyectos", "Seguridad Informatica", "Conocimientos de Negocios", "Normas Estandares", "SAC", "Foros y Comites Externos", "Summit", "Ingles", "Excel", "Otros"`)
    for (var i = 0; i < results.length; i++) {
        let object = results[i];
        writter.write(`"${object.file}", "${object.name}", "${object.fortaleza1}", "${object.fortaleza2}", "${object.fortaleza3}", "${object.enfoque1}", "${object.enfoque2}", "${object.enfoque3}", "${object.plazo1}", "${object.plazo2}", "${object.plazo3}", "${object.otros}", "${object.auditoriaIso}", "${object.certificacionesVisa}", "${object.certificacionesTecnologicas}", "${object.itil}", "${object.pci}", "${object.gestionProyectos}", "${object.seguridadInformatica}", "${object.conocimientoNegocio}", "${object.normasEstandares}", "${object.sac}", "${object.forosYComitesExternos}", "${object.summit}", "${object.ingles}", "${object.excel}", "${object.otrosCursos}"`);
    }
    writter.end();
}

const getFilesToProcess = () => {
    return fs.readdirSync('input');
};

const init = () => {
    let files = getFilesToProcess();
    let results = files.map(file => readFile(`input/${file}`));
    writeResults(results);
}

init();