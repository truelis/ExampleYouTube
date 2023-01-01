from flask import Flask, request, redirect, send_file
import pandas as pd
import io

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Read the uploaded file
        file = request.files["file"]
        df = pd.read_excel(file)

        # Pivot the data and calculate the additional fields
        pivoted = df.pivot_table(index="Day", values=["Impr.", "Clicks", "Conversions", "Cost","All conv. value"], aggfunc="sum")
        pivoted["CTR"] = pivoted["Clicks"] / pivoted["Impr."]
        pivoted["CR"] = pivoted["Conversions"] / pivoted["Clicks"]
        pivoted["ROAS"] = pivoted["All conv. value"] / pivoted["Cost"]
        pivoted["GP"] = pivoted["All conv. value"] - pivoted["Cost"]
        pivoted["Margin"] = (pivoted["All conv. value"] - pivoted["Cost"]) / pivoted["All conv. value"]

        # Create a response containing the updated file
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Original')
        pivoted.to_excel(writer, sheet_name='Analysed')
        workbook  = writer.book
        worksheet = writer.sheets['Analysed']
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({
            'values': '=Analysed!$J$1:$J$70000',
            'categories': '=Analysed!$A$1:$A$70000',
            'name': '=Analysed!$B$1'
        })
        worksheet.insert_chart('E2', chart)
        writer.save()
        output.seek(0)

        # Send the updated file to the user
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            attachment_filename="output.xlsx",
            as_attachment=True
        )

    return """
        <html>
            <body>
                <form method="post" enctype="multipart/form-data">
                    <input type="file" name="file">
                    <input type="submit" value="Upload">
                </form>
            </body>
        </html>
    """

if __name__ == "__main__":
    app.run()
