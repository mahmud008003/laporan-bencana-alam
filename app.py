from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        try:
            # Ambil file Excel dari website
            excel = request.files["excel"]

            if excel.filename == "":
                return "Silakan pilih file Excel terlebih dahulu."

            # Simpan sementara
            excel_path = os.path.join(
                UPLOAD_FOLDER,
                excel.filename
            )

            excel.save(excel_path)

            # Buka template Word
            doc = DocxTemplate("BENTUK LAPORAN.docx")

            # Membaca Excel
            # Sesuaikan dengan struktur Excel milikmu
            df = pd.read_excel(
                excel_path,
                sheet_name="Sheet1 (2)",
                header=3
            )

            # Bersihkan nama kolom
            df.columns = (
                df.columns
                .astype(str)
                .str.strip()
                .str.upper()
            )

            context = {}

            for _, row in df.iterrows():

                kode = row["KODE"]
                persen = row["PERSENTASE"]

                # Lewati pembatas atau baris kosong
                if pd.isna(kode) or pd.isna(persen):
                    continue

                kode = str(kode).strip()

                # Hilangkan .0 jika bilangan bulat
                if (
                    isinstance(persen, float)
                    and persen.is_integer()
                ):
                    persen = int(persen)

                # Tambahkan %
                context[kode] = f"{persen}%"

            # Isi Word
            doc.render(context)

            output = "Laporan_Hasil.docx"

            doc.save(output)

            return send_file(
                output,
                as_attachment=True
            )

        except Exception as e:
            return f"Terjadi kesalahan: {str(e)}"

    return render_template("index.html")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)