import React, { Component } from 'react';
import { render } from 'react-dom';
import Hello from './Hello';
import { Button } from 'primereact/button';
import { saveAs } from 'file-saver';
import { Document, Packer, Paragraph, Table, TableCell, TableRow, Media, Alignment,TextRun, HeadingLevel, SymbolRun, WidthType } from "docx";
import './style.css';
import 'primereact/resources/themes/nova-light/theme.css';
import 'primereact/resources/primereact.min.css';
import 'primeicons/primeicons.css';

interface AppProps { }
interface AppState {
  name: string;
}

class App extends Component<AppProps, AppState> {

  constructor(props) {
    super(props);
    this.state = {
      name: 'React'
    };
  }
  generateDoc = () => {

    const doc = new Document();



    const HEADING_1 = new Paragraph({
      text: "1.	Pfizer Contacts:",
      spacing: {
        after: 300,
      },
      heading: HeadingLevel.HEADING_1,
    });

    const HEADING_2 = new Paragraph({
      text: "2.	Process Overview",
      spacing: {
        before: 300,
        after: 200
      },
      heading: HeadingLevel.HEADING_1,
    });

    const table = new Table({
      width: {
        size: 9500,
        type: WidthType.DXA,
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph({
                children: [new TextRun({ text: "Project Name: ", bold: true, size: 30 }), new TextRun(
                  { text: "click here to enter text", size: 24 })]
              })],
            })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: "Company: ", bold: true, size: 30 }), new TextRun({ text: "click here to enter text", size: 24 })] })],
            })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: "Address: ", bold: true, size: 30 }), new TextRun({ text: "click here to enter text", size: 24 })] })],
            })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: "Client Contact (business): ", bold: true, size: 30 }), new TextRun({ text: "click here to enter text\t", size: 24 }), new TextRun({ text: "Email: ", bold: true, size: 30 }), new TextRun({ text: "click here to enter text\t", size: 24 })] })],
            })
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: "Client Contact (technical): ", bold: true, size: 30 }), new TextRun({ text: "click here to enter text\t", size: 24 }), new TextRun({ text: "Email: ", bold: true, size: 30 }), new TextRun({ text: "click here to enter text", size: 24 })] })],
            })
          ]
        })
      ]
    });
    const paragraph1 = new Paragraph({
      children: [
        new TextRun({ text: "Detailed flow diagram can be found in Attachment 1", size: 24 }),
       
       // new SymbolRun({ char: "0052", symbolfont: "Wingdings 2" }),
      ],
    });
    const paragraph2 = new Paragraph({
      children: [
        new TextRun({ text: "Simple Process Flow Diagram", size: 24,  color: "red"})
       // new SymbolRun({ char: "0052", symbolfont: "Wingdings 2" }),
      ],
    });



    doc.addSection({ children: [HEADING_1, table, HEADING_2, paragraph1,paragraph2] });

    Packer.toBlob(doc).then((blob) => {
      // saveAs from FileSaver will download the file
      saveAs(blob, "sample.docx");
    });

  }


  render() {
    return (
      <div>
        <Button label="Word" className="p-button-rounded" onClick={this.generateDoc} />
      </div>
    );
  }
}

render(<App />, document.getElementById('root'));
