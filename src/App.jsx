import "./App.css";
import { Button } from "@mui/material";
import { useEffect, useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import axios from "axios";
import cover from "./assets/cover.png";

import CircularStatic from "./circularprogress";

import { MuiFileInput } from "mui-file-input";

function App() {
  //const [uploadedFile, setUploadedFile] = useState(null);
  // const [progress, setProgress] = useState(0);
  const [data1, setData1] = useState(null);
  const [data2, setData2] = useState(null);
  const [data3, setData3] = useState(null);
  const [data4, setData4] = useState(null);

  const [data5, setData5] = useState(null);
  const [data6, setData6] = useState(null);
  const [data7, setData7] = useState(null);
  const [data8, setData8] = useState(null);

  const [data9, setData9] = useState(null);
  const [data10, setData10] = useState(null);
  const [data11, setData11] = useState(null);
  const [data12, setData12] = useState(null);

  const [data13, setData13] = useState(null);
  const [data14, setData14] = useState(null);
  const [data15, setData15] = useState(null);
  const [data16, setData16] = useState(null);

  const [data17, setData17] = useState(null);
  const [data18, setData18] = useState(null);
  const [data19, setData19] = useState(null);
  const [data20, setData20] = useState(null);

  const [data21, setData21] = useState(null);
  const [data22, setData22] = useState(null);

  const [data23, setData23] = useState(null);
  const [data24, setData24] = useState(null);

  const [data25, setData25] = useState(null);
  const [data26, setData26] = useState(null);

  const [data27, setData27] = useState(null);
  const [data28, setData28] = useState(null);
  const [data29, setData29] = useState(null);
  const [data30, setData30] = useState(null);

  const [data31, setData31] = useState(null);
  const [data32, setData32] = useState(null);

  const [data33, setData33] = useState(null);
  const [data34, setData34] = useState(null);
  const [data35, setData35] = useState(null);
  const [data36, setData36] = useState(null);

  const [data37, setData37] = useState(null);
  const [data38, setData38] = useState(null);

  const [data39, setData39] = useState(null);
  const [data40, setData40] = useState(null);
  const [data41, setData41] = useState(null);
  const [data42, setData42] = useState(null);

  const [data43, setData43] = useState(null);
  const [data44, setData44] = useState(null);

  const [runtimeHours, setRuntimeHours] = useState(null);
  const [runtimeMin, setRuntimeMin] = useState(null);

  const [fetching, setFetching] = useState(false);
  const [file, setFile] = useState(null);

  // const handleFileUpload = (event) => {
  //   const file = event.target.files[0];
  //   setUploadedFile(file);
  // };
  const handleFileUpload = (newFile) => {
    setFile(newFile);
  };

  useEffect(() => {
    if (file) {
      const fetchData = async () => {
        try {
          setFetching(true);

          // const config = {
          //   onUploadProgress: (progressEvent) => {
          //     const loaded = progressEvent.loaded;
          //     const total = progressEvent.total;
          //     const percentCompleted = Math.round((loaded * 100) / total);
          //     setProgress(percentCompleted);
          //   },
          // };

          //CHP
          const response1 = axios.get(
            "https://script.google.com/macros/s/AKfycbz9u2aJoJzA4ZM4g2t_dGQk3McnyPW9rfUFxxdcNbHuNaXf3HRSyyY4nTwiQkByRpMRLg/exec"
          );
          //RNG
          const response2 = axios.get(
            "https://script.google.com/macros/s/AKfycbytd6cV4qfyiuS6QnQPeFZwqccU2-vqv9tnL4cRsffmLVFQOB1FAO6HZSbVD7PjLp_eLA/exec"
          );
          //PKK
          const response3 = axios.get(
            "https://script.google.com/macros/s/AKfycbzyyAG4ZpI2jdUJ3ld6C-_sAulMwTC7pZYsCqHa-R4lfPZgEt_J7gJFV1LXp3hgAmC-/exec"
          );
          //HUA
          const response4 = axios.get(
            "https://script.google.com/macros/s/AKfycbycKKkH20a8rl8-QWB8CrnM_j8eD6LpadIeRn73mDzL_J3h7PX6u2ohagmt9smYiY1PxQ/exec"
          );
          //TSK
          const response5 = axios.get(
            "https://script.google.com/macros/s/AKfycbyWPDK5oe2RhlSxzDsDhVvooI_N90BmhIFdBengbgO7Aq6ZwAgvqhWdI1EyVfkGl9kW/exec"
          );
          //PHT
          const response6 = axios.get(
            "https://script.google.com/macros/s/AKfycbwJBTdWmfkZs8w3vstxTlvKgVvbfeQ-3PEEgFIrPvlIg9K12E_31BHi3wrVtHZfDbKDUg/exec"
          );
          //KPO
          const response7 = axios.get(
            "https://script.google.com/macros/s/AKfycbyJi6htdI-uKUd5dsbMQrDO9VhDTuGEFWdLurhZONePg9_TcTHzSX9rCFKn65REIKRl-w/exec"
          );
          //KAI
          const response8 = axios.get(
            "https://script.google.com/macros/s/AKfycbzryrF3abgTU_9HllNohAJ9_4bSgkT6QRWasBrqQt0puNbJUGOhAI8xiOMzrUD1w2eN/exec"
          );
          //PBI
          const response9 = axios.get(
            "https://script.google.com/macros/s/AKfycbzv9izLctRhphOIn4OD2HUgQQuttu8lOIfrh6Lod9W5mbtVVlyFQC1cwdP81SEtyH6oXQ/exec"
          );
          //PNP
          const response10 = axios.get(
            "https://script.google.com/macros/s/AKfycbzsnI5ILJLerqWqLhKAPpgESAVDA8SNAnP3p2vMRT06SI-jD6GsP8bO4IB4_oeDRTs8Gw/exec"
          );
          //TAS
          const response11 = axios.get(
            "https://script.google.com/macros/s/AKfycbxRFNFuoNBZj5M5xFX7o0mcraXPKSnGNI5DGHToNy9fAJPSU9_C-KqmdFPwtw1dJaQn/exec"
          );
          //SWI
          const response12 = axios.get(
            "https://script.google.com/macros/s/AKfycbywAXbCkS_7zTLEv2hUoPHs6aI9RCohwnX94kDb4j2fpomHbuqIUghsd6YImESsn7xdWA/exec"
          );
          //LGS
          const response13 = axios.get(
            "https://script.google.com/macros/s/AKfycbyazX9wgJTKbF34n4z3QDM1fPLoRwtLho5RVWaHfg52gx-ewz4HgXzHzV9Bc-h9Fhcwew/exec"
          );
          //KBR
          const response14 = axios.get(
            "https://script.google.com/macros/s/AKfycbwCn90h7bXc7L_L570Fx8k_GyEhSUrzo6eKBSrMn8lwYa42EcF1okSGBYiJrkwcZCT5QQ/exec"
          );

          const [
            result1,
            result2,
            result3,
            result4,
            result5,
            result6,
            result7,
            result8,
            result9,
            result10,
            result11,
            result12,
            result13,
            result14,
          ] = await Promise.all([
            response1,
            response2,
            response3,
            response4,
            response5,
            response6,
            response7,
            response8,
            response9,
            response10,
            response11,
            response12,
            response13,
            response14,
          ]);

          //IRD CHP
          setData1(result1.data.IRD_A_LM);
          setData2(result1.data.IRD_A_CN);
          setData3(result1.data.IRD_B_LM);
          setData4(result1.data.IRD_B_CN);

          //IRD RNG
          setData5(result2.data.IRD_A_LM);
          setData6(result2.data.IRD_A_CN);
          setData7(result2.data.IRD_B_LM);
          setData8(result2.data.IRD_B_CN);

          //UPSRuntime RNG Convert time
          setRuntimeHours(parseInt(result2.data.UPSRuntime / 60));
          setRuntimeMin(String(result2.data.UPSRuntime - 60));

          //IRD PKK
          setData9(result3.data.IRD_A_LM);
          setData10(result3.data.IRD_A_CN);
          setData11(result3.data.IRD_B_LM);
          setData12(result3.data.IRD_B_CN);

          //IRD HUA
          setData13(result4.data.IRD_A_LM);
          setData14(result4.data.IRD_A_CN);
          setData15(result4.data.IRD_B_LM);
          setData16(result4.data.IRD_B_CN);

          //IRD  TSK
          setData17(result5.data.IRD_A_LM);
          setData18(result5.data.IRD_A_CN);
          setData19(result5.data.IRD_B_LM);
          setData20(result5.data.IRD_B_CN);

          //IRD  PHT
          setData21(result6.data.IRD_A_LM);
          setData22(result6.data.IRD_A_CN);

          //IRD  KPO
          setData23(result7.data.IRD_A_LM);
          setData24(result7.data.IRD_A_CN);

          //IRD  KAI
          setData25(result8.data.IRD_A_LM);
          setData26(result8.data.IRD_A_CN);

          //IRD  PBI
          setData27(result9.data.IRD_A_LM);
          setData28(result9.data.IRD_A_CN);
          setData29(result9.data.IRD_B_LM);
          setData30(result9.data.IRD_B_CN);

          //IRD PNP
          setData31(result10.data.IRD_A_LM);
          setData32(result10.data.IRD_A_CN);

          //IRD TAS
          setData33(result11.data.IRD_A_LM);
          setData34(result11.data.IRD_A_CN);
          setData35(result11.data.IRD_B_LM);
          setData36(result11.data.IRD_B_CN);

          //IRD SWI
          setData37(result12.data.IRD_A_LM);
          setData38(result12.data.IRD_A_CN);

          //IRD LGS
          setData39(result13.data.IRD_A_LM);
          setData40(result13.data.IRD_A_CN);
          setData41(result13.data.IRD_B_LM);
          setData42(result13.data.IRD_B_CN);

          //IRD KBR
          setData43(result14.data.IRD_A_LM);
          setData44(result14.data.IRD_A_CN);

          console.log(result1.data.IRD_A_LM);

          setFetching(false);
        } catch (error) {
          console.log(error);
          setFetching(false);
        }
      };

      fetchData();
    }
  }, [file]);

  const handleGenerateExcel = () => {
    if (data43 && data44) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        const worksheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[worksheetName];

        // Write data to specific cells{

        // IRD CHP
        XLSX.utils.sheet_add_json(worksheet, [{ I6: data1 }], {
          skipHeader: true,
          origin: "I6",
        });

        XLSX.utils.sheet_add_json(worksheet, [{ J6: data2 }], {
          skipHeader: true,
          origin: "J6",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ K6: data3 }], {
          skipHeader: true,
          origin: "K6",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ L6: data4 }], {
          skipHeader: true,
          origin: "L6",
        });

        // IRD RNG
        XLSX.utils.sheet_add_json(worksheet, [{ I7: data5 }], {
          skipHeader: true,
          origin: "I7",
        });

        XLSX.utils.sheet_add_json(worksheet, [{ J7: data6 }], {
          skipHeader: true,
          origin: "J7",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ K7: data7 }], {
          skipHeader: true,
          origin: "K7",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ L7: data8 }], {
          skipHeader: true,
          origin: "L7",
        });

        // UPS Runtime RNG
        XLSX.utils.sheet_add_json(worksheet, [{ S7: runtimeHours }], {
          skipHeader: true,
          origin: "S7",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ T7: runtimeMin }], {
          skipHeader: true,
          origin: "T7",
        });

        // IRD PKK
        XLSX.utils.sheet_add_json(worksheet, [{ I12: data9 }], {
          skipHeader: true,
          origin: "I12",
        });

        XLSX.utils.sheet_add_json(worksheet, [{ J12: data10 }], {
          skipHeader: true,
          origin: "J12",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ K12: data11 }], {
          skipHeader: true,
          origin: "K12",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ L12: data12 }], {
          skipHeader: true,
          origin: "L12",
        });

        // IRD HUA
        XLSX.utils.sheet_add_json(worksheet, [{ I16: data13 }], {
          skipHeader: true,
          origin: "I16",
        });

        XLSX.utils.sheet_add_json(worksheet, [{ J16: data14 }], {
          skipHeader: true,
          origin: "J16",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ K16: data15 }], {
          skipHeader: true,
          origin: "K16",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ L16: data16 }], {
          skipHeader: true,
          origin: "L16",
        });

        // IRD TSK
        XLSX.utils.sheet_add_json(worksheet, [{ I8: data17 }], {
          skipHeader: true,
          origin: "I8",
        });

        XLSX.utils.sheet_add_json(worksheet, [{ J8: data18 }], {
          skipHeader: true,
          origin: "J8",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ K8: data19 }], {
          skipHeader: true,
          origin: "K8",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ L8: data20 }], {
          skipHeader: true,
          origin: "L8",
        });

        // IRD PHT
        XLSX.utils.sheet_add_json(worksheet, [{ I9: data21 }], {
          skipHeader: true,
          origin: "I9",
        });

        XLSX.utils.sheet_add_json(worksheet, [{ J9: data22 }], {
          skipHeader: true,
          origin: "J9",
        });

        // IRD KPO
        XLSX.utils.sheet_add_json(worksheet, [{ I10: data23 }], {
          skipHeader: true,
          origin: "I10",
        });

        XLSX.utils.sheet_add_json(worksheet, [{ J10: data24 }], {
          skipHeader: true,
          origin: "J10",
        });

        // IRD KAI
        XLSX.utils.sheet_add_json(worksheet, [{ I11: data25 }], {
          skipHeader: true,
          origin: "I11",
        });

        XLSX.utils.sheet_add_json(worksheet, [{ J11: data26 }], {
          skipHeader: true,
          origin: "J11",
        });

        //IRD PBI
        XLSX.utils.sheet_add_json(worksheet, [{ I15: data27 }], {
          skipHeader: true,
          origin: "I15",
        });

        XLSX.utils.sheet_add_json(worksheet, [{ J15: data28 }], {
          skipHeader: true,
          origin: "J15",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ K15: data29 }], {
          skipHeader: true,
          origin: "K15",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ L15: data30 }], {
          skipHeader: true,
          origin: "L15",
        });

        // IRD PNP
        XLSX.utils.sheet_add_json(worksheet, [{ I18: data31 }], {
          skipHeader: true,
          origin: "I18",
        });

        XLSX.utils.sheet_add_json(worksheet, [{ J18: data32 }], {
          skipHeader: true,
          origin: "J18",
        });

        // IRD TAS
        XLSX.utils.sheet_add_json(worksheet, [{ I13: data33 }], {
          skipHeader: true,
          origin: "I13",
        });

        XLSX.utils.sheet_add_json(worksheet, [{ J13: data34 }], {
          skipHeader: true,
          origin: "J13",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ K13: data35 }], {
          skipHeader: true,
          origin: "K13",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ L13: data36 }], {
          skipHeader: true,
          origin: "L13",
        });

        // IRD SWI
        XLSX.utils.sheet_add_json(worksheet, [{ I17: data37 }], {
          skipHeader: true,
          origin: "I17",
        });

        XLSX.utils.sheet_add_json(worksheet, [{ J17: data38 }], {
          skipHeader: true,
          origin: "J17",
        });

        // IRD LGS
        XLSX.utils.sheet_add_json(worksheet, [{ I14: data39 }], {
          skipHeader: true,
          origin: "I14",
        });

        XLSX.utils.sheet_add_json(worksheet, [{ J14: data40 }], {
          skipHeader: true,
          origin: "J14",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ K14: data41 }], {
          skipHeader: true,
          origin: "K14",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ L14: data42 }], {
          skipHeader: true,
          origin: "L14",
        });

        // IRD KBR
        XLSX.utils.sheet_add_json(worksheet, [{ I19: data43 }], {
          skipHeader: true,
          origin: "I19",
        });

        XLSX.utils.sheet_add_json(worksheet, [{ J19: data44 }], {
          skipHeader: true,
          origin: "J19",
        });

        // Generate Excel file
        const excelBuffer = XLSX.write(workbook, {
          bookType: "xlsx",
          type: "array",
        });
        const blob = new Blob([excelBuffer], {
          type: "application/octet-stream",
        });

        // Save Excel file
        saveAs(blob, "data.xlsx");
      };

      reader.readAsArrayBuffer(file);
    }
  };

  return (
    <>
      <div>
        {/* File upload input */}
        {/* <input type="file" onChange={handleFileUpload} /> */}

        <div className="flex flex-col space-y-5 bo">
          <div className=" md:flex ">
            <img src={cover} className="rounded-2xl" />
          </div>

          <div></div>

          <div className="text-5xl font-extrabold">
            <span className="bg-clip-text text-transparent bg-gradient-to-r from-pink-500 to-violet-500">
              Auto Diary Report 2023
            </span>
          </div>
          <diV></diV>
          <p className="text-2xl">Chumphon Engineering Center </p>
        </div>

        <div className="flex flex-col space-y-10">
          <div></div>
          <MuiFileInput
            className="bg-gradient-to-r from-indigo-500 via-purple-500 to-pink-500"
            value={file}
            onChange={handleFileUpload}
            color="warning"
          />
        </div>

        {/* Display the fetched data or progress indicator */}
        {fetching ? (
          <div style={{ display: "flex", alignItems: "center" }}>
            <CircularStatic />
          </div>
        ) : (
          <>
            {data1 && <div>IRD A LM CHP: {data1}</div>}
            {data2 && <div>IRD A CN CHP: {data2}</div>}
            {data3 && <div>IRD B LM CHP: {data3}</div>}
            {data4 && <div>IRD B CN CHP: {data4}</div>}

            {data5 && <div>IRD A LM RNG: {data5}</div>}
            {data6 && <div>IRD A CN RNG: {data6}</div>}
            {data7 && <div>IRD B LM RNG: {data7}</div>}
            {data8 && <div>IRD B CN RNG: {data8}</div>}
            {runtimeHours && <div>UPS Runtime Hours RNG: {runtimeHours}</div>}
            {runtimeMin && <div>UPS Runtime Min RNG: {runtimeMin}</div>}

            {data9 && <div>IRD A LM PKK: {data9}</div>}
            {data10 && <div>IRD A CN PKK: {data10}</div>}
            {data11 && <div>IRD B LM PKK: {data11}</div>}
            {data12 && <div>IRD B CN PKK: {data12}</div>}

            {data13 && <div>IRD A LM HUA: {data13}</div>}
            {data14 && <div>IRD A CN HUA: {data14}</div>}
            {data15 && <div>IRD B LM HUA: {data15}</div>}
            {data16 && <div>IRD B CN HUA: {data16}</div>}

            {data17 && <div>IRD A LM TSK: {data17}</div>}
            {data18 && <div>IRD A CN TSK: {data18}</div>}
            {data19 && <div>IRD B LM TSK: {data19}</div>}
            {data20 && <div>IRD B CN TSK: {data20}</div>}

            {data21 && <div>IRD A LM PHT: {data21}</div>}
            {data22 && <div>IRD A CN PHT: {data22}</div>}

            {data23 && <div>IRD A LM KPO: {data23}</div>}
            {data24 && <div>IRD A CN KPO: {data24}</div>}

            {data25 && <div>IRD A LM KAI: {data25}</div>}
            {data26 && <div>IRD A CN KAI: {data26}</div>}

            {data27 && <div>IRD A LM PBI: {data27}</div>}
            {data28 && <div>IRD A CN PBI: {data28}</div>}
            {data29 && <div>IRD B LM PBI: {data29}</div>}
            {data30 && <div>IRD B CN PBI: {data30}</div>}

            {data31 && <div>IRD A LM PNP: {data31}</div>}
            {data32 && <div>IRD A CN PNP: {data32}</div>}

            {data33 && <div>IRD A LM TAS: {data33}</div>}
            {data34 && <div>IRD A CN TAS: {data34}</div>}
            {data35 && <div>IRD B LM TAS: {data35}</div>}
            {data36 && <div>IRD B CN TAS: {data36}</div>}

            {data37 && <div>IRD A LM SWI: {data37}</div>}
            {data38 && <div>IRD A CN SWI: {data38}</div>}

            {data39 && <div>IRD A LM LGS: {data39}</div>}
            {data40 && <div>IRD A CN LGS: {data40}</div>}
            {data41 && <div>IRD B LM LGS: {data41}</div>}
            {data42 && <div>IRD B CN LGS: {data42}</div>}

            {data43 && <div>IRD A LM KBR: {data43}</div>}
            {data44 && <div>IRD A CN KBR: {data44}</div>}
          </>
        )}
        {/* Button to trigger Excel file creation */}
        <Button
          variant="contained"
          onClick={handleGenerateExcel}
          disabled={!data43 || !data44 || fetching}
          style={{ marginTop: "16px" }}
        >
          Generate Excel
        </Button>
      </div>
    </>
  );
}

export default App;
