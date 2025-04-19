    let pegawaiIndex = 1;

    document.getElementById("addPegawaiBtn").addEventListener("click", function () {
      const container = document.getElementById("pegawaiContainer");
      const index = pegawaiIndex++;
      const pegawaiHTML = `
        <div class="border p-3 mb-3" id="pegawai_${index}">
          <div class="d-flex justify-content-between">
            <h5>Pegawai ${index + 1}</h5>
            <button type="button" class="btn btn-sm btn-danger" onclick="removePegawai(${index})">Hapus</button>
          </div>
          <div class="row g-3">
            <div class="col-md-6">
              <label class="form-label">Nama</label>
              <input type="text" class="form-control" name="pegawai[${index}][nama]" required />
            </div>
            <div class="col-md-6">
              <label class="form-label">NIP</label>
              <input type="text" class="form-control" name="pegawai[${index}][nip]" />
            </div>
            <div class="col-md-4">
              <label class="form-label">Volume Hari</label>
              <input type="number" class="form-control" name="pegawai[${index}][vol]" required />
            </div>
            <div class="col-md-4">
              <label class="form-label">Nominal Harian</label>
              <div class="input-group">
                <span class="input-group-text">Rp</span>
                <input type="text" class="form-control currency" name="pegawai[${index}][nom]" required />
              </div>
            </div>
          </div>
        </div>
      `;
      container.insertAdjacentHTML("beforeend", pegawaiHTML);
      attachCurrencyFormat();
    });

    function removePegawai(index) {
      const pegawaiDiv = document.getElementById("pegawai_" + index);
      if (pegawaiDiv) pegawaiDiv.remove();
    }

    function formatCurrency(value) {
      const number = parseInt(value.replace(/\D/g, "")) || 0;
      return number.toLocaleString("id-ID");
    }

    function attachCurrencyFormat() {
      document.querySelectorAll("input.currency").forEach((input) => {
        input.addEventListener("input", function (e) {
          const cursor = this.selectionStart;
          this.value = formatCurrency(this.value);
          this.setSelectionRange(cursor, cursor);
        });
      });
    }

    attachCurrencyFormat();

    document.getElementById("dopForm").addEventListener("submit", function (e) {
      e.preventDefault();

      const formData = new FormData(e.target);
      const data = {};
      formData.forEach((value, key) => {
        if (key.includes("[") && key.includes("]")) {
          const match = key.match(/(.*?)\[(\d+)\]\[(.*?)\]/);
          if (match) {
            const [_, prefix, index, field] = match;
            if (!data[prefix]) data[prefix] = [];
            if (!data[prefix][index]) data[prefix][index] = {};
            data[prefix][index][field] = value;
          }
        } else {
          data[key] = value;
        }
      });

      document.getElementById("previewContent").textContent = JSON.stringify(data, null, 2);
      const previewModal = new bootstrap.Modal(document.getElementById("previewModal"));
      previewModal.show();
    });

  async function loadTemplateFromGitHub() {
    const url = "https://raw.githubusercontent.com/kusumaindraputra/doc_generator/main/template/template_doc.docx";
    const response = await fetch(url);
    const arrayBuffer = await response.arrayBuffer();
    return arrayBuffer;
  }

  function generateWordDoc(templateBuffer, data) {
    const zip = new PizZip(templateBuffer);
    const doc = new window.docxtemplater().loadZip(zip);
    doc.setData(data);

    try {
      doc.render();
    } catch (e) {
      console.error(e);
      alert("Gagal menghasilkan dokumen. Periksa template atau data.");
      return;
    }

    const blob = doc.getZip().generate({ type: "blob" });
    const filename = `${(data.tujuan_perjalanan || "dokumen").replace(/\\s+/g, "_")}_${data.tgl_mulai || "tanggal"}.docx`;
    saveAs(blob, filename);
  }

  document.getElementById("downloadBtn").addEventListener("click", async () => {
    const formData = new FormData(document.getElementById("dopForm"));
    const data = {};
    formData.forEach((value, key) => {
      if (key.includes("[")) {
        const match = key.match(/(.*?)\\[(\\d+)\\]\\[(.*?)\\]/);
        if (match) {
          const [, prefix, index, field] = match;
          data[prefix] = data[prefix] || [];
          data[prefix][index] = data[prefix][index] || {};
          data[prefix][index][field] = value;
        }
      } else {
        data[key] = value;
      }
    });

    const buffer = await loadTemplateFromGitHub();
    generateWordDoc(buffer, data);
  });