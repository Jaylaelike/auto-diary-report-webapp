import "./App.css";
import { Button } from "@mui/material";
import { useEffect, useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import axios from "axios";
import cover from "./assets/cover.png";

import Snackbar from "@mui/material/Snackbar";
import MuiAlert from "@mui/material/Alert";
import Table from "@mui/material/Table";
import TableBody from "@mui/material/TableBody";
import TableCell from "@mui/material/TableCell";
import TableContainer from "@mui/material/TableContainer";
import TableHead from "@mui/material/TableHead";
import TableRow from "@mui/material/TableRow";
import Paper from "@mui/material/Paper";

import CircularStatic from "./circularprogress";

import { MuiFileInput } from "mui-file-input";

function App() {
  //const [uploadedFile, setUploadedFile] = useState(null);
  // const [progress, setProgress] = useState(0);
  const [runtimeHours, setRuntimeHours] = useState(null);
  const [runtimeMin, setRuntimeMin] = useState(null);
  const [showAlert, setShowAlert] = useState(false);
  //const[showTable, setShowTable] = useState(false);

  const [fetching, setFetching] = useState(false);
  const [file, setFile] = useState(null);
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

  const [data45, setData45] = useState(null);

  const [data47, setData47] = useState(null);
  // const [data48, setData48] = useState(null);
  const [data49, setData49] = useState(null);
  const [data50, setData50] = useState(null);
  const [data51, setData51] = useState(null);
  const [data52, setData52] = useState(null);
  const [data53, setData53] = useState(null);
  const [data54, setData54] = useState(null);
  const [data55, setData55] = useState(null);
  // const [data56, setData56] = useState(null);
  const [data57, setData57] = useState(null);
  const [data58, setData58] = useState(null);
  const [data59, setData59] = useState(null);
  const [data60, setData60] = useState(null);

  const [data61, setData61] = useState(null);
  const [data62, setData62] = useState(null);
  const [data63, setData63] = useState(null);
  const [data64, setData64] = useState(null);

  const [data65, setData65] = useState(null);
  const [data66, setData66] = useState(null);
  const [data67, setData67] = useState(null);
  const [data68, setData68] = useState(null);

  const [data69, setData69] = useState(null);
  const [data70, setData70] = useState(null);
  const [data71, setData71] = useState(null);
  const [data72, setData72] = useState(null);

  const [data73, setData73] = useState(null);
  const [data74, setData74] = useState(null);
  const [data75, setData75] = useState(null);
  const [data76, setData76] = useState(null);
  const [data77, setData77] = useState(null);
  const [data78, setData78] = useState(null);

  const [data80, setData80] = useState(null);

  // const handleFileUpload = (event) => {
  //   const file = event.target.files[0];
  //   setUploadedFile(file);
  // };
  //handle FielUpload
  const handleFileUpload = (newFile) => {
    setFile(newFile);
  };

  //Create data fot table
  function createData(
    name,
    chp,
    rng,
    tsk,
    pht,
    kpo,
    kai,
    pkk,
    tas,
    lhs,
    pbi,
    hua,
    swi,
    pnp,
    kbr
  ) {
    return {
      name,
      chp,
      rng,
      tsk,
      pht,
      kpo,
      kai,
      pkk,
      tas,
      lhs,
      pbi,
      hua,
      swi,
      pnp,
      kbr,
    };
  }

  const rows = [
    createData(
      "ชุมพร",
      `${data1}`,
      `${data2}`,
      `${data3}`,
      `${data4}`,
      `${data61}`,
      `${data63}`,
      `${data62}`,
      `${data64}`,
      `${data53}`,
      `${data45}`
    ),
    createData(
      "ระนอง",
      `${data5}`,
      `${data6}`,
      `${data7}`,
      `${data8}`,
      `${data66}`,
      `${data67}`,
      `${data65}`,
      `${data68}`,
      `${data54}`,
      `${data80}`
    ),
    createData(
      "ทับสะแก",
      `${data17}`,
      `${data18}`,
      `${data19}`,
      `${data20}`,
      `${data70}`,
      `${data71}`,
      `${data69}`,
      `${data72}`,
      `${data57}`,
      `${data49}`
    ),
    createData(
      "พะโต๊ะ",
      `${data21}`,
      `${data22}`,
      null,
      null,
      `${data73}`,
      `${data74}`,
      null,
      null,
      `${data58}`,
      `${data50}`
    ),
    createData(
      "กะเปอร์",
      `${data23}`,
      `${data24}`,
      null,
      null,
      `${data75}`,
      `${data76}`,
      null,
      null,
      `${data59}`,
      `${data51}`
    ),
    createData(
      "คุระบุรี",
      `${data25}`,
      `${data26}`,
      null,
      null,
      `${data77}`,
      `${data78}`,
      null,
      null,
      `${data60}`,
      `${data52}`
    ),
    createData(
      "ประจวบ(เขาทุ่งกระต่าย)",
      `${data9}`,
      `${data10}`,
      `${data11}`,
      `${data12}`,
      null,
      null,
      null,
      null,
      `${data55}`,
      `${data47}`
    ),
    createData(
      "ท่าแซะ",
      `${data33}`,
      `${data34}`,
      `${data35}`,
      `${data36}`,
      null,
      null,
      null,
      null,
      null,
      null,
    ),
    createData(
      "หลังสวน",
      `${data39}`,
      `${data40}`,
      `${data41}`,
      `${data42}`,
      null,
      null,
      null,
      null,
      null,
      null,
    ),
    createData(
      "เพชรบุรี",
      `${data27}`,
      `${data28}`,
      `${data29}`,
      `${data30}`,
      null,
      null,
      null,
      null,
      null,
      null,
    ),
    createData(
      "หัวหิน",
      `${data13}`,
      `${data14}`,
      `${data15}`,
      `${data16}`,
      null,
      null,
      null,
      null,
      null,
      null,
    ),
    createData(
      "สวี",
      `${data37}`,
      `${data38}`,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null
    ),
    createData(
      "ปราณบุรี",
      `${data31}`,
      `${data32}`,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null
    ),
    createData(
      "กระบุรี",
      `${data43}`,
      `${data44}`,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null
    ),
  ];

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
          //UPS_LOAD
          const response15 = axios.get(
            "https://script.google.com/macros/s/AKfycbwviJ3VJCyy0d1sRmOI0AfGasqkzVlEsRmoInCBwds5pvBStrJY3sE-K0JlN4TdQeg21g/exec"
          );
          //SWR_CHP
          const response16 = axios.get(
            "https://script.google.com/macros/s/AKfycbz2uX3RfbkDNguEoQN3K3aen1aM2t-GrUMEoJT-GvYqffSdVVJ0bwYm5uW5z1SHHFHaeQ/exec"
          );
          //SWR_RNG
          const response17 = axios.get(
            "https://script.google.com/macros/s/AKfycby-tNv2gLU-siv93gCtwU-Oc_HgQQ2UNXA3rhkH8zh-LI9BeyumnHYVp1-HR1VdjPgJ/exec"
          );
          //SWR_TSK
          const response18 = axios.get(
            "https://script.google.com/macros/s/AKfycbz3iuBRqvBcBnh3_imzbW0uefRUYwZM_TBPWYF_1zXo2FMAPYLQfm9mu6XcZbh6E7wp/exec"
          );
          //SWR_PHT
          const response19 = axios.get(
            "https://script.google.com/macros/s/AKfycby3kmOHaBPIBt_t5XwN_SEtKjssXMvlu1UC89bOSGvt1Wl89ORInBgUeAEC1ohAQbMI/exec"
          );
          //SWR_KPO
          const response20 = axios.get(
            "https://script.google.com/macros/s/AKfycbwTY9-Zj4_2JsL7dW2xlwD9Rne2eyfQmcB4ZFZ-Hn1NN4Z-Y9FYtaTGVsxgJfrvTaSq/exec"
          );
          //SWR_KAI
          const response21 = axios.get(
            "https://script.google.com/macros/s/AKfycbyLksPAsU9lmw8XLASwFZWzLCAwk-g_N3z83Ixg1C-fqDUuSDmwCUpVdA2QcCsf4tQH/exec"
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
            result15,
            result16,
            result17,
            result18,
            result19,
            result20,
            result21,
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
            response15,
            response16,
            response17,
            response18,
            response19,
            response20,
            response21,
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
          setData80(result2.data.UPSRuntime);

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

          //UPS
          setData45(result1.data.UPSRuntime); //CHP
          setData47(result3.data.UPSRuntime); //PKK
          setData49(result5.data.UPSRuntime); //TSK
          setData50(result6.data.UPS_Runtime); //PHT
          setData51(result7.data.UPS_Runtime); //KPO
          setData52(result8.data.UPS_Runtime); //KAI

          //UPS LOAD
          setData53(result15.data.loadCHP); //CHP
          setData54(result15.data.loadRNG); //RNG
          setData55(result15.data.loadPKK); //PKK
          setData57(result15.data.loadTSK); //TSK
          setData58(result15.data.loadPHT); //PHT
          setData59(result15.data.loadKPO); //KPO
          setData60(result15.data.loadKAI); //KAI

          //CHP_SWR
          setData61(result16.data.FWD_upper);
          setData62(result16.data.FWD_lower);
          setData63(result16.data.lower);
          setData64(result16.data.upper);

          //RNG_SWR
          setData65(result17.data.FWD_upper);
          setData66(result17.data.FWD_lower);
          setData67(result17.data.lower);
          setData68(result17.data.upper);

          //TSK_SWR
          setData69(result18.data.FWD_upper);
          setData70(result18.data.FWD_lower);
          setData71(result18.data.lower);
          setData72(result18.data.upper);

          //PHT-SWR
          setData73(parseInt(result19.data.FWD_ant / 1000));
          setData74(result19.data.SWR_ant);

          //KPO-SWR
          setData75(result20.data.FWD_ant);
          setData76(result20.data.SWR_ant);

          //KAI_SWR
          setData77(result21.data.FWD_ant);
          setData78(result21.data.SWR_ant);

          console.log(result21.data.IRD_A_LM);

          setFetching(false);
          setShowAlert(true);
        } catch (error) {
          console.log(error);
          setFetching(false);
        }
      };

      fetchData();
    }
  }, [file]);

  //State close of alert popup
  const handleAlertClose = () => {
    setShowAlert(false);
  };

  const handleTable = () => {
    if (data77 && data78) {
      return tableFunc();
    }
  };

  const handleGenerateExcel = () => {
    if (data77 && data78) {
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

        // UPS RUNTIME CHP
        XLSX.utils.sheet_add_json(worksheet, [{ T6: data45 }], {
          skipHeader: true,
          origin: "T6",
        });

        // UPS RUNTIME TSK
        XLSX.utils.sheet_add_json(worksheet, [{ T8: data49 }], {
          skipHeader: true,
          origin: "T8",
        });

        // UPS RUNTIME PHT
        XLSX.utils.sheet_add_json(worksheet, [{ T9: data50 }], {
          skipHeader: true,
          origin: "T9",
        });

        // UPS RUNTIME KPO
        XLSX.utils.sheet_add_json(worksheet, [{ T10: data51 }], {
          skipHeader: true,
          origin: "T10",
        });

        // UPS RUNTIME KAI
        XLSX.utils.sheet_add_json(worksheet, [{ T11: data52 }], {
          skipHeader: true,
          origin: "T11",
        });

        // UPS RUNTIME PKK
        XLSX.utils.sheet_add_json(worksheet, [{ T12: data47 }], {
          skipHeader: true,
          origin: "T12",
        });

        // UPS LOAD CHP
        XLSX.utils.sheet_add_json(worksheet, [{ R6: data53 }], {
          skipHeader: true,
          origin: "R6",
        });

        // UPS LOAD RNG
        XLSX.utils.sheet_add_json(worksheet, [{ R7: data54 }], {
          skipHeader: true,
          origin: "R7",
        });

        // UPS LOAD TSK
        XLSX.utils.sheet_add_json(worksheet, [{ R8: data57 }], {
          skipHeader: true,
          origin: "R8",
        });

        // UPS LOAD PHT
        XLSX.utils.sheet_add_json(worksheet, [{ R9: data58 }], {
          skipHeader: true,
          origin: "R9",
        });

        // UPS LOAD KPO
        XLSX.utils.sheet_add_json(worksheet, [{ R10: data59 }], {
          skipHeader: true,
          origin: "R10",
        });

        // UPS LOAD KAI
        XLSX.utils.sheet_add_json(worksheet, [{ R11: data60 }], {
          skipHeader: true,
          origin: "R11",
        });

        // UPS LOAD PKK
        XLSX.utils.sheet_add_json(worksheet, [{ R12: data55 }], {
          skipHeader: true,
          origin: "R12",
        });

        // SWR_CHP
        XLSX.utils.sheet_add_json(worksheet, [{ N6: data62 }], {
          skipHeader: true,
          origin: "N6",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ O6: data63 }], {
          skipHeader: true,
          origin: "O6",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ P6: data61 }], {
          skipHeader: true,
          origin: "P6",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ Q6: data64 }], {
          skipHeader: true,
          origin: "Q6",
        });

        // SWR_RNG
        XLSX.utils.sheet_add_json(worksheet, [{ N7: data66 }], {
          skipHeader: true,
          origin: "N7",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ O7: data67 }], {
          skipHeader: true,
          origin: "O7",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ P7: data65 }], {
          skipHeader: true,
          origin: "P7",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ Q7: data68 }], {
          skipHeader: true,
          origin: "Q7",
        });

        // SWR_TSK
        XLSX.utils.sheet_add_json(worksheet, [{ N8: data70 }], {
          skipHeader: true,
          origin: "N8",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ O8: data71 }], {
          skipHeader: true,
          origin: "O8",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ P8: data69 }], {
          skipHeader: true,
          origin: "P8",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ Q8: data72 }], {
          skipHeader: true,
          origin: "Q8",
        });

        // SWR_PHT
        XLSX.utils.sheet_add_json(worksheet, [{ P9: data73 }], {
          skipHeader: true,
          origin: "P9",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ Q9: data74 }], {
          skipHeader: true,
          origin: "Q9",
        });

        // SWR_KPO
        XLSX.utils.sheet_add_json(worksheet, [{ P10: data75 }], {
          skipHeader: true,
          origin: "P10",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ Q10: data76 }], {
          skipHeader: true,
          origin: "Q10",
        });

        // SWR_KAI
        XLSX.utils.sheet_add_json(worksheet, [{ P11: data77 }], {
          skipHeader: true,
          origin: "P11",
        });
        XLSX.utils.sheet_add_json(worksheet, [{ Q11: data78 }], {
          skipHeader: true,
          origin: "Q11",
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
      <div className="flex justify-center">
        <div className="py-12">
          <img src={cover} className="rounded-2xl" />
        </div>
      </div>
      <div className=" grid grid-cols-1 gap-4 content-center">
        {/* File upload input */}
        {/* <input type="file" onChange={handleFileUpload} /> */}

        <div className="flex flex-col space-y-5 ">
          <div className="text-5xl font-extrabold">
            <span className="bg-clip-text text-transparent bg-gradient-to-r from-pink-500 to-violet-500">
              Auto Diary Report 2023
            </span>
          </div>

          <p className="text-2xl bg-clip-text text-transparent bg-gradient-to-r from-pink-500 to-violet-500">
            Chumphon Engineering Center{" "}
          </p>
        </div>

        <div className="flex flex-col space-y-10">
          <MuiFileInput
            className="bg-gradient-to-r from-indigo-500 via-purple-500 to-pink-500"
            value={file}
            onChange={handleFileUpload}
            color="warning"
          />
          <div></div>
        </div>

        {/* Display the fetched data or progress indicator */}
        {fetching ? (
          <div className="justify-center">
            <CircularStatic />
          </div>
        ) : (
          handleTable()
        )}

        {/* Button to trigger Excel file creation */}
        <div className="justify-center">
          <Button
            disableElevation
            onClick={handleGenerateExcel}
            disabled={!data77 || !data78 || fetching}
            style={{ marginTop: "16px" }}
          >
            Generate Excel
          </Button>
        </div>
      </div>

      <Snackbar
        open={showAlert}
        autoHideDuration={3000} // Duration to show the alert in milliseconds (adjust as needed)
        onClose={handleAlertClose}
        anchorOrigin={{
          vertical: "top",
          horizontal: "right",
        }}
      >
        <MuiAlert
          onClose={handleAlertClose}
          severity="success"
          elevation={6}
          variant="filled"
        >
          เย้ 🥳 ทำงานสำเร็จแล้ว โหลดไฟล์ด่านล่างได้เลยคะ 🙏🏻
        </MuiAlert>
      </Snackbar>
    </>
  );

  //Create Table

  function tableFunc() {
    return (
      <div className="grid justify-items-stretch">
        <div>
          <TableContainer component={Paper}>
            <Table sx={{ minWidth: 650 }} aria-label="simple table">
              <TableHead>
                <TableRow>
                  <TableCell>สถานี</TableCell>
                  <TableCell align="right">IRD C/N_A (dB)</TableCell>
                  <TableCell align="right">IRD LM_A (dB)</TableCell>
                  <TableCell align="right">IRD C/N_B (dB)</TableCell>
                  <TableCell align="right">IRD LM_B (dB)</TableCell>
                  <TableCell align="right">FW_LOWER (W)</TableCell>
                  <TableCell align="right">FW_SWR (W)</TableCell>
                  <TableCell align="right">FW_UPPER (W)</TableCell>
                  <TableCell align="right">FW_SWR (W)</TableCell>
                  <TableCell align="right">UPS LOAD (%)</TableCell>
                  <TableCell align="right">UPS Runtime (min)</TableCell>
                  <TableCell align="right"></TableCell>
                </TableRow>
              </TableHead>
              <TableBody>
                {rows.map((row) => (
                  <TableRow
                    key={row.name}
                    sx={{
                      "&:last-child td, &:last-child th": { border: 0 },
                    }}
                  >
                    <TableCell component="th" scope="row">
                      {row.name}
                    </TableCell>
                    <TableCell align="right">{row.chp}</TableCell>
                    <TableCell align="right">{row.rng}</TableCell>
                    <TableCell align="right">{row.tsk}</TableCell>
                    <TableCell align="right">{row.pht}</TableCell>
                    <TableCell align="right">{row.kpo}</TableCell>
                    <TableCell align="right">{row.kai}</TableCell>
                    <TableCell align="right">{row.pkk}</TableCell>
                    <TableCell align="right">{row.tas}</TableCell>
                    <TableCell align="right">{row.lhs}</TableCell>
                    <TableCell align="right">{row.pbi}</TableCell>
                    <TableCell align="right">{row.hua}</TableCell>
                    <TableCell align="right">{row.swi}</TableCell>
                    <TableCell align="right">{row.pnp}</TableCell>
                    <TableCell align="right">{row.kbr}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </TableContainer>
        </div>
      </div>
    );
  }
}

export default App;

// {data1 && <div>IRD A LM CHP: {data1}</div>}
// {data2 && <div>IRD A CN CHP: {data2}</div>}
// {data3 && <div>IRD B LM CHP: {data3}</div>}
// {data4 && <div>IRD B CN CHP: {data4}</div>}
// {data5 && <div>IRD A LM RNG: {data5}</div>}
// {data6 && <div>IRD A CN RNG: {data6}</div>}
// {data7 && <div>IRD B LM RNG: {data7}</div>}
// {data8 && <div>IRD B CN RNG: {data8}</div>}
// {runtimeHours && <div>UPS Runtime Hours RNG: {runtimeHours}</div>}
// {runtimeMin && <div>UPS Runtime Min RNG: {runtimeMin}</div>}
// {data9 && <div>IRD A LM PKK: {data9}</div>}
// {data10 && <div>IRD A CN PKK: {data10}</div>}
// {data11 && <div>IRD B LM PKK: {data11}</div>}
// {data12 && <div>IRD B CN PKK: {data12}</div>}
// {data13 && <div>IRD A LM HUA: {data13}</div>}
// {data14 && <div>IRD A CN HUA: {data14}</div>}
// {data15 && <div>IRD B LM HUA: {data15}</div>}
// {data16 && <div>IRD B CN HUA: {data16}</div>}
// {data17 && <div>IRD A LM TSK: {data17}</div>}
// {data18 && <div>IRD A CN TSK: {data18}</div>}
// {data19 && <div>IRD B LM TSK: {data19}</div>}
// {data20 && <div>IRD B CN TSK: {data20}</div>}
// {data21 && <div>IRD A LM PHT: {data21}</div>}
// {data22 && <div>IRD A CN PHT: {data22}</div>}
// {data23 && <div>IRD A LM KPO: {data23}</div>}
// {data24 && <div>IRD A CN KPO: {data24}</div>}
// {data25 && <div>IRD A LM KAI: {data25}</div>}
// {data26 && <div>IRD A CN KAI: {data26}</div>}
// {data27 && <div>IRD A LM PBI: {data27}</div>}
// {data28 && <div>IRD A CN PBI: {data28}</div>}
// {data29 && <div>IRD B LM PBI: {data29}</div>}
// {data30 && <div>IRD B CN PBI: {data30}</div>}
// {data31 && <div>IRD A LM PNP: {data31}</div>}
// {data32 && <div>IRD A CN PNP: {data32}</div>}
// {data33 && <div>IRD A LM TAS: {data33}</div>}
// {data34 && <div>IRD A CN TAS: {data34}</div>}
// {data35 && <div>IRD B LM TAS: {data35}</div>}
// {data36 && <div>IRD B CN TAS: {data36}</div>}
// {data37 && <div>IRD A LM SWI: {data37}</div>}
// {data38 && <div>IRD A CN SWI: {data38}</div>}
// {data39 && <div>IRD A LM LGS: {data39}</div>}
// {data40 && <div>IRD A CN LGS: {data40}</div>}
// {data41 && <div>IRD B LM LGS: {data41}</div>}
// {data42 && <div>IRD B CN LGS: {data42}</div>}
// {data43 && <div>IRD A LM KBR: {data43}</div>}
// {data44 && <div>IRD A CN KBR: {data44}</div>}
// {data45 && <div>UPS RUNTIME CHP: {data45}</div>}

// {data47 && <div>UPS RUNTIME PKK: {data47}</div>}
// {data49 && <div>UPS RUNTIME TSK: {data49}</div>}
// {data50 && <div>UPS RUNTIME PHT: {data50}</div>}
// {data51 && <div>UPS RUNTIME KPO: {data51}</div>}
// {data52 && <div>UPS RUNTIME KAI: {data52}</div>}
// {data53 && <div>UPS LOAD CHP: {data53}</div>}
// {data54 && <div>UPS LOAD RNG: {data54}</div>}
// {data55 && <div>UPS LOAD PKK: {data55}</div>}
// {data57 && <div>UPS LOAD TSK: {data57}</div>}
// {data58 && <div>UPS LOAD PHT: {data58}</div>}
// {data59 && <div>UPS LOAD KPO: {data59}</div>}
// {data60 && <div>UPS LOAD KAI: {data60}</div>}
