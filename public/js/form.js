let excelColumns = [];
let docxFields = [];
let workbook = null;
let xlsxFile = null;
let docxFile = null;

document.getElementById('fileInput').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array' });

            const sheetSelect = document.getElementById('sheetSelect');
            sheetSelect.innerHTML = '';
            workbook.SheetNames.forEach(sheetName => {
                const option = document.createElement('option');
                option.value = sheetName;
                option.innerText = sheetName;
                sheetSelect.appendChild(option);
            });

            sheetSelect.disabled = false;
            document.getElementById('loadFiles').disabled = false;
        };
        reader.readAsArrayBuffer(file);
    }
});

document.getElementById('fileForm').addEventListener('submit', async (event) => {
    event.preventDefault();

    xlsxFile = document.getElementById('fileInput').files[0];
    docxFile = document.getElementById('docxInput').files[0];
    const selectedSheet = document.getElementById('sheetSelect').value;

    if (workbook && docxFile && selectedSheet) {
        const worksheet = workbook.Sheets[selectedSheet];
        const excelData = XLSX.utils.sheet_to_json(worksheet);

        docxFields = await readDOCX(docxFile);
        excelColumns = Object.keys(excelData[0]);
        generateMappingForm(docxFields, excelColumns);
        generateNamingForm(excelColumns);
    } else {
        alert('Veuillez sélectionner un fichier DOCX et une feuille Excel.');
    }
});

function readDOCX(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const arrayBuffer = e.target.result;
            mammoth.extractRawText({ arrayBuffer: arrayBuffer })
                .then(result => {
                    const textContent = result.value;
                    const matches = textContent.match(/<<\s*([^>]+)\s*>>/g);
                    resolve(matches || []);
                })
                .catch(err => {
                    console.error('Erreur lors de la lecture du fichier DOCX:', err);
                    resolve([]);
                });
        };
        reader.readAsArrayBuffer(file);
    });
}

function generateMappingForm(docxFields, excelColumns) {
    const container = document.getElementById('fieldsContainer');
    container.innerHTML = '';

    if (docxFields.length === 0) {
        container.innerHTML = '<p>Aucun champ trouvé dans le fichier DOCX.</p>';
        return;
    }

    docxFields.forEach((field, index) => {
        const fieldRow = document.createElement('div');
        fieldRow.className = 'field-row';

        const label = document.createElement('label');
        label.innerText = `Champ DOCX : ${field}`;
        label.setAttribute('for', `mapping-${index}`);

        const select = document.createElement('select');
        select.id = `mapping-${index}`;
        select.name = field;
        select.required = true;

        const emptyOption = document.createElement('option');
        emptyOption.value = '';
        emptyOption.innerText = 'Sélectionner une colonne Excel';
        select.appendChild(emptyOption);

        excelColumns.forEach(column => {
            const option = document.createElement('option');
            option.value = column;
            option.innerText = column;
            select.appendChild(option);
        });

        fieldRow.appendChild(label);
        fieldRow.appendChild(select);
        container.appendChild(fieldRow);
    });
}

function generateNamingForm(excelColumns) {
    const selectSortBy = document.getElementById('sortBy');
    selectSortBy.innerHTML = '';
    selectSortBy.disabled = false;
    selectSortBy.required = true;

    const selectIdNameOptions = document.getElementById('idNamingOptions');
    selectIdNameOptions.innerHTML = '';
    selectIdNameOptions.disabled = false;
    selectIdNameOptions.required = true;

    excelColumns.forEach(column => {
        const option = document.createElement('option');
        option.value = column;
        option.innerText = column;
        selectSortBy.appendChild(option);
        selectIdNameOptions.appendChild(option.cloneNode(true));
    });
}

document.getElementById('autoMap').addEventListener('click', () => {
    const selects = document.querySelectorAll('#fieldsContainer select');

    selects.forEach(select => {
        const docxField = select.name.replace(/<<\s*|\s*>>/g, '').toLowerCase();
        let bestMatch = null;
        let bestScore = Infinity;

        excelColumns.forEach(column => {
            const columnNormalized = column.toLowerCase();
            const score = levenshteinDistance(docxField, columnNormalized);

            if (score < bestScore) {
                bestScore = score;
                bestMatch = column;
            }
        });

        if (bestMatch) {
            select.value = bestMatch;
        }
    });
    alert('Correspondances automatiques suggérées.');
});

function levenshteinDistance(a, b) {
    const matrix = [];

    for (let i = 0; i <= a.length; i++) {
        matrix[i] = [i];
    }
    for (let j = 0; j <= b.length; j++) {
        matrix[0][j] = j;
    }

    for (let i = 1; i <= a.length; i++) {
        for (let j = 1; j <= b.length; j++) {
            if (a[i - 1] === b[j - 1]) {
                matrix[i][j] = matrix[i - 1][j - 1];
            } else {
                matrix[i][j] = Math.min(
                    matrix[i - 1][j] + 1,
                    matrix[i][j - 1] + 1,
                    matrix[i - 1][j - 1] + 1
                );
            }
        }
    }

    return matrix[a.length][b.length];
}

document.getElementById('generateQuotes').addEventListener('click', () => {
    const mappingForm = document.getElementById('mappingForm');
    const mappingData = {};
    const mappingInputs = mappingForm.querySelectorAll('select');

    mappingInputs.forEach(input => {
        if (input.value) {
            const fieldName = input.name.replace(/<<\s*|\s*>>/g, ''); // Remove chevrons
            mappingData[fieldName] = input.value;
        }
    });

    const namingForm = document.getElementById('namingForm');
    const namingData = {};
    const namingInputs = namingForm.querySelectorAll('input, select');

    namingInputs.forEach(input => {
        if (input.value && input.name) {
            namingData[input.name] = input.value;
        }
    });

    namingData.sortBy = document.getElementById('sortBy').value;
    namingData.idNamingOptions = document.getElementById('idNamingOptions').value;
    namingData.sheetSelect = document.getElementById('sheetSelect').value;

    const formData = new FormData();
    formData.append('xlsxFile', xlsxFile);
    formData.append('docxFile', docxFile);
    formData.append('mapping', JSON.stringify(mappingData));
    formData.append('naming', JSON.stringify(namingData));

    console.log('FormData structure:', formData);

    fetch('http://localhost:3000/generate-quotes', {
        method: 'POST',
        body: formData
    })
        .then(response => {
            if (!response.ok) {
                throw new Error('Erreur lors de la création des devis.');
            }
            return response.json();
        })
        .then(data => {
            // alert('Devis créés avec succès !');
            console.log(data);
            document.getElementById('downloadZip').disabled = false;
            document.getElementById('downloadZip').dataset.downloadCode = data.downloadCode;
        })
        .catch(error => {
            console.error('Erreur :', error);
            alert('Une erreur est survenue. Veuillez réessayer.');
        });
});

document.getElementById('downloadZip').addEventListener('click', () => {
    const downloadCode = document.getElementById('downloadZip').dataset.downloadCode;
    if (downloadCode) {
        window.location.href = `http://localhost:3000/download-zip/${downloadCode}`;
    }
});