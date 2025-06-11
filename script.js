let tableData = [];

document.getElementById('upload').addEventListener('change', function (e) {
  const file = e.target.files[0];
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    tableData = json.map(row => ({
      fornecimento: row["Fornecimento"],
      awb: row["AWB coleta"]
    })).filter(r => r.fornecimento && r.awb);

    renderTable(tableData);
  };
  reader.readAsArrayBuffer(file);
});

function renderTable(data) {
  const tbody = document.querySelector("#dataTable tbody");
  tbody.innerHTML = "";
  data.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row.fornecimento}</td>
      <td>${row.awb}</td>
      <td><button class="print-btn" onclick="generateEtiqueta('${row.fornecimento}', '${row.awb}')">🖨️</button></td>
    `;
    tbody.appendChild(tr);
  });
}

function generateEtiqueta(fornecimento, awb) {
  // Remove any existing print iframe
  const oldIframe = document.getElementById('print-iframe');
  if (oldIframe) oldIframe.remove();

  // Create a hidden iframe
  const iframe = document.createElement('iframe');
iframe.style.display = 'none';
  iframe.id = 'print-iframe';
  document.body.appendChild(iframe);

const htmlContent = `
    <html>
    <head>
        <title>Etiqueta</title>
        <style>
            @page {
                size: 100mm 40mm;
                margin: 0;
            }
            body {
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 5mm;
                display: flex;
                align-items: center;
                justify-content: space-between;
                height: 40mm;
                box-sizing: border-box;
            }
            .text {
                font-size: 14pt;
                display: flex;
                flex-direction: column;
                justify-content: center;
            }
            .text b {
                font-size: 16pt;
            }
            .qrcode {
                margin-left: 10px;
            }
        </style>
    </head>
    <body>
        <div class="text">
            <b>Remessa:</b>
            ${fornecimento}<br>
            <b>AWB:</b>
            ${awb}
        </div>
        <div class="qrcode" id="qrcode"></div>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
        <script>
            window.onload = function () {
                var qrcodeDiv = document.getElementById("qrcode");
                if (qrcodeDiv) {
                    new QRCode(qrcodeDiv, {
                        text: "${awb}",
                        width: 80,
                        height: 80
                    });
                    setTimeout(function () {
                        window.focus();
                        window.print();
                    }, 500);
                }
            };
        </script>
    </body>
    </html>
`;

  iframe.contentWindow.document.open();
  iframe.contentWindow.document.write(htmlContent);
  iframe.contentWindow.document.close();

  // Remove the iframe after printing
  iframe.contentWindow.onafterprint = function () {
    setTimeout(() => {
      iframe.remove();
    }, 500);
  };
}

document.getElementById('search').addEventListener('input', function () {
  const searchTerm = this.value.toLowerCase();
  const filtered = tableData.filter(item =>
    item.fornecimento.toString().toLowerCase().includes(searchTerm) ||
    item.awb.toString().toLowerCase().includes(searchTerm)
  );
  renderTable(filtered);
});
