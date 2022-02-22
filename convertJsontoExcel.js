 function export2Csv(JSONData, FileTitle, ShowLabel) {
    //If JSONData is not an object then JSON.parse will parse the JSON string in an Object
    var arrData = typeof JSONData != "object" ? JSON.parse(JSONData) : JSONData;
    var CSV = "";
    //This condition will generate the Label/Header
    if (ShowLabel) {
      var row = "";
      //This loop will extract the label from 1st index of on array
      for (var colobj of arrData.Columns.Column) {
        //Now convert each value to string and comma-seprated
        row += colobj.Name + ",";
      }
      //this removes the comma
      row = row.slice(0, -1);
      //append Label row with line break
      CSV += row + "\r\n";
    }
    //1st loop is to extract each row
    for (var i = 0; i < arrData.Row.length; i++) {
      var row = "";
      //2nd loop will extract each column and convert it in string comma-seprated

      for (var key in arrData.Row[i]) {
        row += '"' + arrData.Row[i][key] + '",';
      }
      //row.slice(0, row.length - 1);
      //add a line break after each row
      CSV += row + "\r\n";
    }

    if (CSV == "") {
      alert("Invalid data");
      return;
    }
    //Generate a file name
    var filename = FileTitle;
    var blob = new Blob([CSV], {
      type: "text/csv;charset=utf-8;",
    });
    if (navigator.msSaveBlob) {
      // IE 10+
      navigator.msSaveBlob(blob, filename);
    } else {
      var link = document.createElement("a");
      if (link.download !== undefined) {
        // feature detection
        // Browsers that support HTML5 download attribute
        var url = URL.createObjectURL(blob);
        link.setAttribute("href", url);
        link.style = "visibility:hidden";
        link.download = filename + ".csv";
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
      }
    }
  }

  function export2Xml(data, fileName) {
    let root = `<?xml version="1.0"?>
    <?mso-application progid="Excel.Sheet"?>`;

    let work_book = `
    <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
     xmlns:o="urn:schemas-microsoft-com:office:office"
     xmlns:x="urn:schemas-microsoft-com:office:excel"
     xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882"
     xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
     xmlns:html="http://www.w3.org/TR/REC-html40">`;

    let styles = `<Styles>
    <Style ss:ID="Default" ss:Name="Normal">
     <Alignment ss:Vertical="Bottom"/>
     <Borders/>
     <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
     <Interior/>
     <NumberFormat/>
     <Protection/>
    </Style>
    <Style ss:ID="border-color">
      <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
      </Borders>
      <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#FFFFFF"/>
      <Interior ss:Color="#70AD47" ss:Pattern="Solid"/>
      <NumberFormat/>
    </Style>
    <Style ss:ID="border">
    <Borders>
      <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
      <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
    </Borders>
  </Style>
   </Styles>`;
    work_book += styles;

    let createWorkSheet = (name, sheet_idx, data) => {
      let work_sheet = `<Worksheet ss:Name="${name}">
          <Table ss:DefaultRowHeight="14.5">`;
      let body = "";

      data.forEach((row, idx) => {
        body += `<Row ss:AutoFitHeight="0">`;
        if (idx == 0) {
          row.forEach((cell, cell_idx) => {
            body += `<Cell ss:StyleID="border-color">
                <Data ss:Type="String">${cell}</Data></Cell>`;
          });
        } else {
          row.forEach((cell) => {
            body += `<Cell ss:StyleID="border"><Data ss:Type="String">${cell}</Data></Cell>`;
          });
        }

        body += "</Row>";
      });

      work_sheet += body + `</Table></Worksheet>`;
      return work_sheet;
    };

    data.forEach((worksheet, idx) => {
      work_book += createWorkSheet(worksheet.name, idx + 1, worksheet.data);
    });

    work_book += `</Workbook>`;
    root += work_book;
    //Download file
    downloadfile(root, fileName, "text/xml");
  }

  function downloadfile(fileString, filename, content_type) {
    var blob = new Blob([fileString], {
      type: `${content_type};charset=utf-8;`,
    });
    var link = document.createElement("a");
    if (link.download !== undefined) {
      var url = URL.createObjectURL(blob);
      link.setAttribute("href", url);
      link.style = "visibility:hidden";
      link.download = filename + ".xml";
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  }

  export {export2Csv,export2Xml}
