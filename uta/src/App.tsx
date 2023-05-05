import { useState, useCallback } from 'react'
import './App.css'
import { useDropzone } from 'react-dropzone';
import * as XLSX from 'xlsx';

function App() {
  const [workbook, setWorkbook] = useState({data: null, cols: null});
  const [fileName, setFileName] = useState(null);
  const [excelData, setExcelData] = useState(null);
  const [chartData, setChartData] = useState(null);

  const onDrop = useCallback((acceptedFiles: any) => {
    console.log(acceptedFiles);
    console.log(acceptedFiles[0].path);
    console.log(acceptedFiles[0].name);

    const reader = new FileReader();
    const rABS = !!reader.readAsBinaryString;
    reader.onload = e => {
      /* Parse data */
      const bstr = e.target?.result;
      const workbook = XLSX.read(bstr, { type: rABS ? "binary" : "array" });

      /* Get first worksheet */
      const wsname = workbook.SheetNames[0];
      const ws = workbook.Sheets[wsname];
      console.log(rABS, workbook);
      /* Convert array of arrays */
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
      /* Update state */
      setWorkbook({ data: data as any, cols: make_cols(ws["!ref"] as string) as any });
    };
    if (rABS) reader.readAsBinaryString(acceptedFiles[0]);
    else reader.readAsArrayBuffer(acceptedFiles[0]);
  }, []);

  const make_cols = (refstr: string) => {
    const o = [];
    const C = XLSX.utils.decode_range(refstr).e.c + 1;
    for (let i = 0; i < C; ++i) o[i] = { name: XLSX.utils.encode_col(i), key: i };
    return o;
  };

  const { getRootProps, getInputProps, isDragActive } = useDropzone({ onDrop });

  const filterData = () => {
    console.log("WIP");
  }


  return (
    <div>
      <div {...getRootProps()}>
        <input {...getInputProps()} />
          {isDragActive ? (
            <p>Drop the files here ...</p>
          ) : (
            <p>Drag 'n' drop some files here, or click to select files</p>
          )}
        </div>
        <table className="table table-striped">
          <thead>
            <tr>
              {workbook.cols !== null ? (workbook.cols.map((c: any) => (
                <th key={c.key}>{c.name}</th>
              ))) : (<span>no item</span>)}
            </tr>
          </thead>
          <tbody>
            {workbook.data?.map((r: any, i: any) => (
              <tr key={i}>
                {i}
                {workbook.cols?.map((c: any) => (
                  <td key={c.key}>{r[c.key]}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
    {chartData && (
      <div>
        {/* <Bar data={chartData} /> */}
      </div>
    )}
  </div>
  )
}

export default App
