import "@fontsource/montserrat/400.css";
import "@fontsource/montserrat/700.css";

import "@fontsource/noto-sans/400.css";
import "@fontsource/noto-sans/700.css";

import "@fontsource/noto-serif/400.css";
import "@fontsource/noto-serif/700.css";

import "@fontsource/playfair-display/400.css";
import "@fontsource/playfair-display/700.css";

import "@fontsource/cormorant-garamond/400.css";
import "@fontsource/cormorant-garamond/700.css";

import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import QRCode from "qrcode";
import jsPDF from "jspdf";

const FONT_URL =
  "https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;500;600;700&family=Inter:wght@400;500;600;700;800&family=Lora:wght@400;500;600;700&family=Merriweather:wght@400;700&family=Montserrat:wght@400;500;600;700;800&family=Nunito:wght@400;500;600;700;800&family=Open+Sans:wght@400;600;700;800&family=Oswald:wght@400;500;600;700&family=PT+Sans:wght@400;700&family=PT+Serif:wght@400;700&family=Playfair+Display:wght@400;500;600;700;800&family=Raleway:wght@400;500;600;700;800&display=swap";

const fontLink = document.createElement("link");
fontLink.rel = "stylesheet";
fontLink.href = FONT_URL;

if (!document.head.querySelector(`link[href="${FONT_URL}"]`)) {
  document.head.appendChild(fontLink);
}

function App() {
  // =========================================================
  // FILES
  // =========================================================

  const [excelFile, setExcelFile] = useState(null);
  const [designFile, setDesignFile] = useState(null);

  const [students, setStudents] = useState([]);
  const [designUrl, setDesignUrl] = useState("");

  // =========================================================
  // A4
  // =========================================================

  const [orientation, setOrientation] = useState("landscape");

  const canvasSize = useMemo(() => {
    if (orientation === "portrait") {
      return {
        width: 794,
        height: 1123,
      };
    }

    return {
      width: 1123,
      height: 794,
    };
  }, [orientation]);

  // =========================================================
  // ELEMENTS
  // =========================================================

  const [elements, setElements] = useState([
    {
      id: "type",
      label: "Түрі",
      x: 561,
      y: 150,
      fontFamily: "Arial",
      fontSize: 42,
      fontWeight: "bold",
      fontStyle: "normal",
      color: "#000000",
      textAlign: "center",
      uppercase: false,
      visible: true,
    },

    {
      id: "name",
      label: "Оқушының аты-жөні",
      x: 561,
      y: 300,
      fontFamily: "Arial",
      fontSize: 30,
      fontWeight: "bold",
      fontStyle: "normal",
      color: "#000000",
      textAlign: "center",
      uppercase: false,
      visible: true,
    },

    {
      id: "district",
      label: "Аудан",
      x: 561,
      y: 360,
      fontFamily: "Arial",
      fontSize: 20,
      fontWeight: "normal",
      fontStyle: "normal",
      color: "#000000",
      textAlign: "center",
      uppercase: false,
      visible: true,
    },

    {
      id: "institution",
      label: "Мекеме атауы",
      x: 561,
      y: 405,
      fontFamily: "Arial",
      fontSize: 20,
      fontWeight: "normal",
      fontStyle: "normal",
      color: "#000000",
      textAlign: "center",
      uppercase: false,
      visible: true,
    },

    {
      id: "leader",
      label: "Жетекшісінің аты-жөні",
      x: 561,
      y: 450,
      fontFamily: "Arial",
      fontSize: 18,
      fontWeight: "normal",
      fontStyle: "normal",
      color: "#000000",
      textAlign: "center",
      uppercase: false,
      visible: true,
    },

    {
      id: "subject",
      label: "Пәні",
      x: 561,
      y: 495,
      fontFamily: "Arial",
      fontSize: 20,
      fontWeight: "normal",
      fontStyle: "normal",
      color: "#000000",
      textAlign: "center",
      uppercase: false,
      visible: true,
    },

    {
      id: "nomination",
      label: "Номинация",
      x: 561,
      y: 520,
      fontFamily: "Arial",
      fontSize: 20,
      fontWeight: "normal",
      fontStyle: "normal",
      color: "#000000",
      textAlign: "center",
      uppercase: false,
      visible: true,
    },

    {
      id: "place",
      label: "Жүлделі орын",
      x: 561,
      y: 550,
      fontFamily: "Arial",
      fontSize: 25,
      fontWeight: "bold",
      fontStyle: "normal",
      color: "#000000",
      textAlign: "center",
      uppercase: false,
      visible: true,
    },

    {
      id: "registration",
      label: "Тіркеу №",
      x: 180,
      y: 700,
      fontFamily: "Arial",
      fontSize: 18,
      fontWeight: "normal",
      fontStyle: "normal",
      color: "#000000",
      textAlign: "left",
      uppercase: false,
      visible: true,
    },

    {
      id: "qr",
      label: "QR код",
      x: 950,
      y: 680,
      fontFamily: "Arial",
      fontSize: 18,
      fontWeight: "normal",
      fontStyle: "normal",
      color: "#000000",
      textAlign: "center",
      uppercase: false,
      visible: true,
      qrSize: 300,
    },
  ]);

  const [selectedElement, setSelectedElement] = useState("name");
  const [dragging, setDragging] = useState(null);

  // =========================================================
  // PDF
  // =========================================================

  const [pdfLoading, setPdfLoading] = useState(false);
  const [pdfProgress, setPdfProgress] = useState("");
  const [errorMessage, setErrorMessage] = useState("");

  // =========================================================
  // PREVIEW
  // =========================================================

  const [qrPreview, setQrPreview] = useState("");

  const editorRef = useRef(null);

  // =========================================================
  // EXCEL COLUMNS
  // =========================================================

  const columnMap = {
    district: "Аудан",
    institution: "Мекеме атауы",
    leader: "Жетекшісінің аты-жөні",
    name: "Оқушының аты-жөні",
    registration: "Тіркеу №",
    competition: "Байқау атауы",
    subject: "Пәні",
    nomination: "Номинация",
    type: "Түрі",
    place: "Жүлделі орын",
    order: "Өткізу бұйрық номері/күні",
  };

  // =========================================================
  // SAFE VALUE
  // =========================================================

  const normalizeHeader = (value) => {
    return String(value ?? "")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  };

  const findColumn = (student, wanted) => {
    if (!student) return "";

    const keys = Object.keys(student);

    const exact = keys.find(
      (key) => normalizeHeader(key) === normalizeHeader(wanted)
    );

    if (exact) {
      return student[exact] ?? "";
    }

    return "";
  };

  const getStudentValue = (student, id) => {
    if (!student) return "";

    const column = columnMap[id];

    return findColumn(student, column);
  };

  // =========================================================
  // EXCEL
  // =========================================================

  const handleExcelUpload = (event) => {
    try {
      const file = event.target.files?.[0];

      if (!file) return;

      setExcelFile(file);
      setErrorMessage("");
      setPdfProgress("");

      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const buffer = e.target.result;

          const workbook = XLSX.read(buffer, {
            type: "array",
          });

          if (!workbook.SheetNames.length) {
            throw new Error("Excel ішінде парақ табылмады.");
          }

          const sheet = workbook.Sheets[workbook.SheetNames[0]];

          const rows = XLSX.utils.sheet_to_json(sheet, {
            defval: "",
            raw: false,
          });

          if (!rows.length) {
            throw new Error("Excel файлы бос.");
          }

          setStudents(rows);

          console.log("Excel оқылды:", rows.length);
        } catch (error) {
          console.error(error);

          setErrorMessage(
            "Excel оқу кезінде қате: " +
              (error?.message || String(error))
          );
        }
      };

      reader.onerror = () => {
        setErrorMessage("Excel файлын оқу мүмкін болмады.");
      };

      reader.readAsArrayBuffer(file);
    } catch (error) {
      console.error(error);

      setErrorMessage(
        "Excel қате: " + (error?.message || String(error))
      );
    }
  };

  // =========================================================
  // DESIGN
  // =========================================================

  const handleDesignUpload = (event) => {
    try {
      const file = event.target.files?.[0];

      if (!file) return;

      setDesignFile(file);
      setErrorMessage("");

      if (designUrl) {
        URL.revokeObjectURL(designUrl);
      }

      const url = URL.createObjectURL(file);

      setDesignUrl(url);
    } catch (error) {
      console.error(error);

      setErrorMessage("Дизайн жүктеу кезінде қате болды.");
    }
  };

  // =========================================================
  // STUDENT PREVIEW
  // =========================================================

  const firstStudent = students.length ? students[0] : null;

  // =========================================================
  // TEXT
  // =========================================================

  const getElementText = (element, student) => {
    if (!student || element.id === "qr") {
      return element.id === "qr" ? "QR" : "";
    }

    let text = String(getStudentValue(student, element.id) ?? "");

    if (element.uppercase) {
      text = text.toUpperCase();
    }

    return text;
  };

  // =========================================================
  // QR TEXT
  // =========================================================

  const getQrText = (student) => {
    if (!student) return "";

    const fields = [
      ["Аудан", "district"],
      ["Мекеме атауы", "institution"],
      ["Жетекшісі", "leader"],
      ["Оқушы", "name"],
      ["Тіркеу №", "registration"],
      ["Байқау", "competition"],
      ["Пәні", "subject"],
      ["Номинация", "nomination"],
      ["Түрі", "type"],
      ["Жүлделі орын", "place"],
      ["Бұйрық", "order"],
    ];

    const qrLines = ["ДИПЛОМ ТУРАЛЫ АҚПАРАТ"];

    fields.forEach(([label, id]) => {
      const value = String(
        getStudentValue(student, id) ?? ""
      ).trim();

      if (value !== "") {
        qrLines.push(`${label}: ${value}`);
      }
    });

    return qrLines.join("\n");
  };

  // =========================================================
  // QR PREVIEW
  // =========================================================

  useEffect(() => {
    let cancelled = false;

    const createPreview = async () => {
      if (!firstStudent) {
        setQrPreview("");
        return;
      }

      try {
        const qr = await QRCode.toDataURL(
          getQrText(firstStudent),
          {
            width: 500,
            margin: 1,
            errorCorrectionLevel: "H",
            color: {
              dark: "#000000",
              light: "#ffffff",
            },
          }
        );

        if (!cancelled) {
          setQrPreview(qr);
        }
      } catch (error) {
        console.error("QR preview error:", error);
      }
    };

    createPreview();

    return () => {
      cancelled = true;
    };
  }, [students]);

  // =========================================================
  // ELEMENT UPDATE
  // =========================================================

  const updateElement = (property, value) => {
    setElements((current) =>
      current.map((element) => {
        if (element.id !== selectedElement) {
          return element;
        }

        return {
          ...element,
          [property]: value,
        };
      })
    );
  };

  const currentElement =
    elements.find(
      (element) => element.id === selectedElement
    ) || elements[0];

  // =========================================================
  // CENTER
  // =========================================================

  const centerSelectedElement = () => {
    setElements((current) =>
      current.map((element) => {
        if (element.id !== selectedElement) {
          return element;
        }

        return {
          ...element,
          x: Math.round(canvasSize.width / 2),
          y: Math.round(canvasSize.height / 2),
        };
      })
    );
  };

  // =========================================================
  // DRAG
  // =========================================================

  const startDrag = (event, id) => {
    event.preventDefault();
    event.stopPropagation();

    if (!editorRef.current) return;

    const rect = editorRef.current.getBoundingClientRect();

    const scaleX = canvasSize.width / rect.width;
    const scaleY = canvasSize.height / rect.height;

    const element = elements.find(
      (item) => item.id === id
    );

    if (!element) return;

    const mouseX =
      (event.clientX - rect.left) * scaleX;

    const mouseY =
      (event.clientY - rect.top) * scaleY;

    setSelectedElement(id);

    setDragging({
      id,
      offsetX: mouseX - element.x,
      offsetY: mouseY - element.y,
    });
  };

  const moveDrag = (event) => {
    if (!dragging || !editorRef.current) return;

    const rect = editorRef.current.getBoundingClientRect();

    const scaleX = canvasSize.width / rect.width;
    const scaleY = canvasSize.height / rect.height;

    let x =
      (event.clientX - rect.left) * scaleX -
      dragging.offsetX;

    let y =
      (event.clientY - rect.top) * scaleY -
      dragging.offsetY;

    x = Math.max(
      0,
      Math.min(canvasSize.width, x)
    );

    y = Math.max(
      0,
      Math.min(canvasSize.height, y)
    );

    setElements((current) =>
      current.map((element) => {
        if (element.id !== dragging.id) {
          return element;
        }

        return {
          ...element,
          x: Math.round(x),
          y: Math.round(y),
        };
      })
    );
  };

  const stopDrag = () => {
    setDragging(null);
  };

  // =========================================================
  // KEYBOARD
  // =========================================================

  useEffect(() => {
    const handleKeyDown = (event) => {
      if (!selectedElement) return;

      const keys = [
        "ArrowUp",
        "ArrowDown",
        "ArrowLeft",
        "ArrowRight",
      ];

      if (!keys.includes(event.key)) return;

      event.preventDefault();

      const step = event.shiftKey ? 10 : 1;

      setElements((current) =>
        current.map((element) => {
          if (element.id !== selectedElement) {
            return element;
          }

          let x = element.x;
          let y = element.y;

          if (event.key === "ArrowLeft") x -= step;
          if (event.key === "ArrowRight") x += step;
          if (event.key === "ArrowUp") y -= step;
          if (event.key === "ArrowDown") y += step;

          x = Math.max(
            0,
            Math.min(canvasSize.width, x)
          );

          y = Math.max(
            0,
            Math.min(canvasSize.height, y)
          );

          return {
            ...element,
            x,
            y,
          };
        })
      );
    };

    window.addEventListener(
      "keydown",
      handleKeyDown
    );

    return () => {
      window.removeEventListener(
        "keydown",
        handleKeyDown
      );
    };
  }, [
    selectedElement,
    canvasSize.width,
    canvasSize.height,
  ]);

  // =========================================================
  // IMAGE LOADER
  // =========================================================

  const loadImage = (src) => {
    return new Promise((resolve, reject) => {
      const img = new Image();

      img.onload = () => resolve(img);

      img.onerror = () => {
        reject(
          new Error(
            "Диплом фонының суретін жүктеу мүмкін болмады."
          )
        );
      };

      img.src = src;
    });
  };

  // =========================================================
  // QR IMAGE LOADER
  // =========================================================

  const loadQrImage = (src) => {
    return new Promise((resolve, reject) => {
      const img = new Image();

      img.onload = () => resolve(img);

      img.onerror = () => {
        reject(
          new Error("QR суретін жүктеу мүмкін болмады.")
        );
      };

      img.src = src;
    });
  };

  // =========================================================
  // CANVAS TEXT
  // =========================================================

  const drawTextElement = (
    ctx,
    element,
    student
  ) => {
    const text = getElementText(
      element,
      student
    );

    if (!text) return;

    ctx.save();

    const weight =
      element.fontWeight === "bold"
        ? "bold"
        : "normal";

    const style =
      element.fontStyle === "italic"
        ? "italic"
        : "normal";

    ctx.font =
      `${style} ${weight} ${element.fontSize}px "${element.fontFamily}"`;

    ctx.fillStyle = element.color;

    ctx.textAlign =
      element.textAlign || "center";

    ctx.textBaseline = "middle";

    ctx.fillText(
      text,
      element.x,
      element.y
    );

    ctx.restore();
  };

  // =========================================================
  // CREATE ONE DIPLOMA CANVAS — HIGH QUALITY
  // =========================================================

  const createDiplomaCanvas = async (
    student,
    backgroundImage
  ) => {
    const SCALE = 4;

    const canvas = document.createElement("canvas");

    canvas.width = canvasSize.width * SCALE;
    canvas.height = canvasSize.height * SCALE;

    const ctx = canvas.getContext("2d", {
      alpha: false,
      willReadFrequently: false,
    });

    if (!ctx) {
      throw new Error(
        "Canvas құрылғысын жасау мүмкін болмады."
      );
    }

    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = "high";

    // Ақ фон
    ctx.fillStyle = "#ffffff";

    ctx.fillRect(
      0,
      0,
      canvas.width,
      canvas.height
    );

    // Диплом фоны
    ctx.drawImage(
      backgroundImage,
      0,
      0,
      canvas.width,
      canvas.height
    );

    // =======================================================
    // QR
    // =======================================================

    let qrImage = null;

    const qrElement = elements.find(
      (element) => element.id === "qr"
    );

    if (
      qrElement &&
      qrElement.visible !== false
    ) {
      const qrText = getQrText(student);

      if (qrText) {
        const qrDataUrl =
          await QRCode.toDataURL(
            qrText,
            {
              width: 1800,
              margin: 1,
              errorCorrectionLevel: "H",
              color: {
                dark: "#000000",
                light: "#ffffff",
              },
            }
          );

        qrImage =
          await loadQrImage(qrDataUrl);
      }
    }

    // =======================================================
    // ELEMENTS
    // =======================================================

    for (const element of elements) {
      if (element.visible === false) {
        continue;
      }

      // Түрі және Жүлделі орын диплом бетіне шықпайды.
      // Бірақ QR ішінде қалады.
      if (
        element.id === "type" ||
        element.id === "place"
      ) {
        continue;
      }

      // =====================================================
      // QR
      // =====================================================

      if (element.id === "qr") {
        if (!qrImage) continue;

        const size =
          (Number(element.qrSize) || 120) *
          SCALE;

        const x = element.x * SCALE;
        const y = element.y * SCALE;

        ctx.save();

        // QR нүктелері бұлыңғыр болмауы үшін
        ctx.imageSmoothingEnabled = false;

        ctx.drawImage(
          qrImage,
          x - size / 2,
          y - size / 2,
          size,
          size
        );

        ctx.restore();

        continue;
      }

      // =====================================================
      // TEXT
      // =====================================================

      const text =
        getElementText(
          element,
          student
        );

      if (!text) continue;

      ctx.save();

      const weight =
        element.fontWeight === "bold"
          ? "bold"
          : "normal";

      const style =
        element.fontStyle === "italic"
          ? "italic"
          : "normal";

      ctx.font =
        `${style} ${weight} ${element.fontSize * SCALE}px "${element.fontFamily}"`;

      ctx.fillStyle = element.color;

      ctx.textAlign =
        element.textAlign || "center";

      ctx.textBaseline = "middle";

      ctx.fillText(
        text,
        element.x * SCALE,
        element.y * SCALE
      );

      ctx.restore();
    }

    return canvas;
  };

  // =========================================================
  // PDF
  // =========================================================

  const generateAllPDF = async () => {
    if (pdfLoading) return;

    setErrorMessage("");
    setPdfProgress("");

    try {
      // Excel тексеру
      if (!students.length) {
        throw new Error(
          "Алдымен Excel файлын жүктеңіз."
        );
      }

      // Дизайн тексеру
      if (!designUrl) {
        throw new Error(
          "Алдымен диплом дизайнын жүктеңіз."
        );
      }

      setPdfLoading(true);

      // Шрифттер дайын болғанша күтеміз
      await document.fonts.ready;

      setPdfProgress(
        "⏳ Диплом фоны дайындалып жатыр..."
      );

      const backgroundImage =
        await loadImage(designUrl);

      const isLandscape =
        orientation === "landscape";

      const pdf =
        new jsPDF({
          orientation: isLandscape
            ? "landscape"
            : "portrait",
          unit: "mm",
          format: "a4",
          compress: true,
        });

      const pageWidth =
        isLandscape ? 297 : 210;

      const pageHeight =
        isLandscape ? 210 : 297;

      // =====================================================
      // ӘР ОҚУШЫҒА БІР БЕТ
      // =====================================================

      for (
        let i = 0;
        i < students.length;
        i++
      ) {
        const student =
          students[i];

        setPdfProgress(
          `⏳ ${i + 1} / ${students.length} диплом дайындалып жатыр...`
        );

        const diplomaCanvas =
          await createDiplomaCanvas(
            student,
            backgroundImage
          );

        /*
          PNG қолданамыз.
          Бұл QR кодтың сапасын жақсы сақтайды.
        */
        const imageData =
          diplomaCanvas.toDataURL(
            "image/png"
          );

        // Бірінші беттен кейін жаңа бет
        if (i > 0) {
          pdf.addPage(
            "a4",
            isLandscape
              ? "landscape"
              : "portrait"
          );
        }

        /*
          Нақты PNG ретінде саламыз.
        */
        pdf.addImage(
          imageData,
          "PNG",
          0,
          0,
          pageWidth,
          pageHeight,
          undefined,
          "FAST"
        );

        // Canvas жадысын босату
        diplomaCanvas.width = 1;
        diplomaCanvas.height = 1;

        /*
          Браузерге демалу үшін
          әр 5 беттен кейін кішкене кідіріс.
        */
        if (
          i % 5 === 0 &&
          i !== 0
        ) {
          await new Promise(
            (resolve) =>
              setTimeout(
                resolve,
                50
              )
          );
        }
      }

      setPdfProgress(
        "📄 PDF дайын! Файл сақталып жатыр..."
      );

      const fileName =
        `Дипломдар_${students.length}_бет.pdf`;

      pdf.save(fileName);

      setPdfProgress(
        `✅ ДАЙЫН! ${students.length} диплом бір PDF файлына жасалды.`
      );
    } catch (error) {
      console.error(
        "PDF ERROR:",
        error
      );

      let message =
        "Белгісіз қате.";

      if (
        error &&
        typeof error.message ===
          "string"
      ) {
        message =
          error.message;
      } else if (
        typeof error === "string"
      ) {
        message = error;
      }

      setErrorMessage(
        "❌ PDF жасау кезінде қате болды: " +
          message
      );

      setPdfProgress("");
    } finally {
      setPdfLoading(false);
    }
  };

  // =========================================================
  // RENDER ELEMENT
  // =========================================================

  const renderPreviewElement = (
    element
  ) => {
    const text =
      getElementText(
        element,
        firstStudent
      );

    if (
      element.id === "qr"
    ) {
      return qrPreview ? (
        <img
          src={qrPreview}
          alt="QR"
          draggable="false"
          className="preview-qr"
        />
      ) : (
        <div className="qr-empty">
          QR
        </div>
      );
    }

    return (
      text ||
      element.label
    );
  };

  // =========================================================
  // RENDER
  // =========================================================

  return (
    <div
      className="app"
      onMouseMove={moveDrag}
      onMouseUp={stopDrag}
      onMouseLeave={stopDrag}
    >
      {/* ===================================================
          TOP
      =================================================== */}

      <header className="topbar">
        <div>
          <h1>
            🏆 Диплом генераторы
          </h1>

          <div className="subtitle">
            Excel → Дизайн → QR → Бір PDF
          </div>
        </div>

        <div className="student-count">
          {students.length
            ? `👨‍🎓 ${students.length} оқушы`
            : "Excel күтілуде"}
        </div>
      </header>

      <main className="main-container">

        {/* =================================================
            WORKSPACE
        ================================================= */}

        <section className="workspace">

          {/* FILES */}

          <div className="steps-card">

            <div className="step-box">
              <div className="step-number">
                1
              </div>

              <div className="step-content">
                <b>
                  Excel файлы
                </b>

                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={
                    handleExcelUpload
                  }
                />

                {excelFile && (
                  <small className="success">
                    ✅ {excelFile.name}
                  </small>
                )}
              </div>
            </div>

            <div className="step-box">
              <div className="step-number">
                2
              </div>

              <div className="step-content">
                <b>
                  Диплом дизайны
                </b>

                <input
                  type="file"
                  accept=".png,.jpg,.jpeg"
                  onChange={
                    handleDesignUpload
                  }
                />

                {designFile && (
                  <small className="success">
                    ✅ {designFile.name}
                  </small>
                )}
              </div>
            </div>

            <div className="step-box">
              <div className="step-number">
                3
              </div>

              <div className="step-content">
                <b>
                  A4 бағыты
                </b>

                <div className="orientation-buttons">

                  <button
                    className={
                      orientation ===
                      "landscape"
                        ? "active"
                        : ""
                    }
                    onClick={() =>
                      setOrientation(
                        "landscape"
                      )
                    }
                  >
                    ▭ Альбом
                  </button>

                  <button
                    className={
                      orientation ===
                      "portrait"
                        ? "active"
                        : ""
                    }
                    onClick={() =>
                      setOrientation(
                        "portrait"
                      )
                    }
                  >
                    ▯ Кітапша
                  </button>

                </div>
              </div>
            </div>

          </div>

          {/* EDITOR */}

          {designUrl ? (
            <div className="editor-card">

              <div className="editor-header">

                <div>
                  <h2>
                    🎨 Диплом макеті
                  </h2>

                  <p>
                    Мәтінді мышкамен
                    жылжыт. Таңдалған
                    элементті стрелкамен
                    дәлде.
                  </p>
                </div>

                <div className="a4-badge">
                  A4 •{" "}
                  {orientation ===
                  "landscape"
                    ? "Альбом"
                    : "Кітапша"}
                </div>

              </div>

              {/* EDITOR STAGE */}

              <div className="editor-stage">

                <div
                  ref={editorRef}
                  className="editor"
                  style={{
                    width:
                      canvasSize.width,
                    height:
                      canvasSize.height,
                    aspectRatio:
                      `${canvasSize.width}/${canvasSize.height}`,
                  }}
                >

                  {/* BACKGROUND */}

                  <img
                    src={designUrl}
                    alt="Диплом"
                    className="design-image"
                    draggable="false"
                  />

                  {/* CENTER CROSS */}

                  <div className="center-line-x" />
                  <div className="center-line-y" />

                  <div className="center-cross">
                    +
                  </div>

                  {/* ELEMENTS */}

                  {elements.map(
                    (element) => {

                      if (
                        element.visible ===
                        false
                      ) {
                        return null;
                      }

                      const selected =
                        selectedElement ===
                        element.id;

                      return (
                        <div
                          key={
                            element.id
                          }
                          className={
                            "design-element " +
                            (
                              selected
                                ? "selected-element"
                                : ""
                            ) +
                            (
                              element.id ===
                              "qr"
                                ? " qr-element"
                                : ""
                            )
                          }
                          style={{
                            left:
                              element.x,
                            top:
                              element.y,
                            fontFamily:
                              element.fontFamily,
                            fontSize:
                              element.fontSize,
                            fontWeight:
                              element.fontWeight,
                            fontStyle:
                              element.fontStyle,
                            color:
                              element.color,
                            textAlign:
                              element.textAlign,
                            width:
                              element.id ===
                              "qr"
                                ? element.qrSize
                                : "auto",
                            height:
                              element.id ===
                              "qr"
                                ? element.qrSize
                                : "auto",
                          }}
                          onMouseDown={(
                            e
                          ) =>
                            startDrag(
                              e,
                              element.id
                            )
                          }
                          onClick={() =>
                            setSelectedElement(
                              element.id
                            )
                          }
                        >

                          {renderPreviewElement(
                            element
                          )}

                          {selected && (
                            <span className="element-cross">
                              +
                            </span>
                          )}

                        </div>
                      );
                    }
                  )}

                </div>

              </div>

              <div className="editor-help">

                <span>
                  🖱️ Мышка — жылжыту
                </span>

                <span>
                  ⬅️⬆️⬇️➡️ — 1 px
                </span>

                <span>
                  Shift + стрелка — 10 px
                </span>

                <span>
                  🎯 Крестик — центр
                </span>

              </div>

            </div>
          ) : (

            <div className="empty-editor">

              <div className="empty-icon">
                🖼️
              </div>

              <h2>
                Диплом дизайнын
                жүктеңіз
              </h2>

              <p>
                PNG немесе JPG
                файлды таңдаңыз.
              </p>

            </div>

          )}

          {/* ERROR */}

          {errorMessage && (
            <div className="error-box">
              {errorMessage}
            </div>
          )}

          {/* PDF */}

          {students.length > 0 &&
            designUrl && (

              <div className="pdf-card">

                <div>

                  <h2>
                    📄 Барлық дипломды
                    жасау
                  </h2>

                  <p>
                    {students.length}{" "}
                    оқушы →{" "}
                    {students.length}{" "}
                    A4 бет → бір PDF
                  </p>

                </div>

                <button
                  className="pdf-button"
                  onClick={
                    generateAllPDF
                  }
                  disabled={
                    pdfLoading
                  }
                >
                  {pdfLoading
                    ? "⏳ Жасалып жатыр..."
                    : "📥 БАРЛЫҚ ДИПЛОМДЫ PDF ЖАСАУ"}
                </button>

                {pdfProgress && (
                  <div className="pdf-progress">
                    {pdfProgress}
                  </div>
                )}

              </div>

            )}

          {/* EXCEL TABLE */}

          {students.length > 0 && (

            <div className="table-card">

              <div className="table-header">

                <div>

                  <h2>
                    📊 Excel мәліметтері
                  </h2>

                  <span>
                    Барлығы:{" "}
                    <b>
                      {
                        students.length
                      }
                    </b>{" "}
                    оқушы
                  </span>

                </div>

              </div>

              <div className="table-wrapper">

                <table>

                  <thead>

                    <tr>

                      {Object.keys(
                        students[0]
                      ).map(
                        (column) => (
                          <th
                            key={
                              column
                            }
                          >
                            {
                              column
                            }
                          </th>
                        )
                      )}

                    </tr>

                  </thead>

                  <tbody>

                    {students
                      .slice(
                        0,
                        10
                      )
                      .map(
                        (
                          row,
                          index
                        ) => (

                          <tr
                            key={
                              index
                            }
                          >

                            {Object.keys(
                              students[0]
                            ).map(
                              (
                                column
                              ) => (

                                <td
                                  key={
                                    column
                                  }
                                >
                                  {
                                    row[
                                      column
                                    ]
                                  }
                                </td>

                              )
                            )}

                          </tr>

                        )
                      )}

                  </tbody>

                </table>

              </div>

              {students.length >
                10 && (

                <div className="table-more">

                  Тағы{" "}
                  {students.length -
                    10}{" "}
                  оқушы бар.
                  PDF жасағанда
                  барлығы өңделеді.

                </div>

              )}

            </div>

          )}

        </section>

        {/* =================================================
            SETTINGS
        ================================================= */}

        <aside className="settings-panel">

          <div className="settings-title">

            <div>

              <h2>
                ⚙️ Баптау
              </h2>

              <span>
                Таңдалған элемент
              </span>

            </div>

            <div className="element-icon">
              ✦
            </div>

          </div>

          {/* ELEMENT SELECT */}

          <div className="element-list">

            {elements.map(
              (element) => (

                <button
                  key={
                    element.id
                  }
                  className={
                    selectedElement ===
                    element.id
                      ? "element-select active"
                      : "element-select"
                  }
                  onClick={() =>
                    setSelectedElement(
                      element.id
                    )
                  }
                >
                  {
                    element.label
                  }
                </button>

              )
            )}

          </div>

          <div className="selected-name">
            {
              currentElement?.label
            }
          </div>

          {/* X Y */}

          <div className="coordinates">

            <div className="coordinate">

              <span>
                X
              </span>

              <input
                type="number"
                value={Math.round(
                  currentElement?.x ||
                    0
                )}
                onChange={(e) =>
                  updateElement(
                    "x",
                    Number(
                      e.target.value
                    )
                  )
                }
              />

            </div>

            <div className="coordinate">

              <span>
                Y
              </span>

              <input
                type="number"
                value={Math.round(
                  currentElement?.y ||
                    0
                )}
                onChange={(e) =>
                  updateElement(
                    "y",
                    Number(
                      e.target.value
                    )
                  )
                }
              />

            </div>

          </div>

          <button
            className="center-button"
            onClick={
              centerSelectedElement
            }
          >
            🎯 Дәл ортасына қою
          </button>

          {/* FONT */}

          {currentElement?.id !==
            "qr" && (
            <>

              {/* FONT */}

              <div className="setting-group">

                <label>
                  Шрифт
                </label>

                <select
                  value={
                    currentElement?.fontFamily
                  }
                  onChange={(e) =>
                    updateElement(
                      "fontFamily",
                      e.target.value
                    )
                  }
                >

                  <option value="Arial">
                    Arial
                  </option>

                  <option value="Times New Roman">
                    Times New Roman
                  </option>

                  <option value="Georgia">
                    Georgia
                  </option>

                  <option value="Verdana">
                    Verdana
                  </option>

                  <option value="Tahoma">
                    Tahoma
                  </option>

                  <option value="Trebuchet MS">
                    Trebuchet MS
                  </option>

                  <option value="Montserrat">
                    Montserrat
                  </option>

                  <option value="Playfair Display">
                    Playfair Display
                  </option>

                  <option value="Merriweather">
                    Merriweather
                  </option>

                  <option value="Lora">
                    Lora
                  </option>

                  <option value="Roboto">
                    Roboto
                  </option>

                  <option value="Open Sans">
                    Open Sans
                  </option>

                  <option value="Noto Sans">
                    Noto Sans
                  </option>

                  <option value="Noto Serif">
                    Noto Serif
                  </option>

                  <option value="PT Sans">
                    PT Sans
                  </option>

                  <option value="PT Serif">
                    PT Serif
                  </option>

                  <option value="Cormorant Garamond">
                    Cormorant Garamond
                  </option>

                  <option value="Raleway">
                    Raleway
                  </option>

                  <option value="Oswald">
                    Oswald
                  </option>

                  <option value="Nunito">
                    Nunito
                  </option>

                </select>

              </div>

              {/* SIZE */}

              <div className="setting-group">

                <div className="label-line">

                  <label>
                    Өлшем
                  </label>

                  <b>
                    {
                      currentElement?.fontSize
                    }
                    px
                  </b>

                </div>

                <input
                  type="range"
                  min="8"
                  max="100"
                  value={
                    currentElement?.fontSize
                  }
                  onChange={(e) =>
                    updateElement(
                      "fontSize",
                      Number(
                        e.target.value
                      )
                    )
                  }
                />

              </div>

              {/* COLOR */}

              <div className="setting-group">

                <label>
                  Мәтін түсі
                </label>

                <div className="color-row">

                  <input
                    type="color"
                    value={
                      currentElement?.color
                    }
                    onChange={(e) =>
                      updateElement(
                        "color",
                        e.target.value
                      )
                    }
                  />

                  <span>
                    {
                      currentElement?.color
                    }
                  </span>

                </div>

              </div>

              {/* B I */}

              <div className="setting-group">

                <label>
                  Стиль
                </label>

                <div className="style-buttons">

                  <button
                    className={
                      currentElement?.fontWeight ===
                      "bold"
                        ? "active"
                        : ""
                    }
                    onClick={() =>
                      updateElement(
                        "fontWeight",
                        currentElement.fontWeight ===
                          "bold"
                          ? "normal"
                          : "bold"
                      )
                    }
                  >
                    <b>
                      B
                    </b>
                  </button>

                  <button
                    className={
                      currentElement?.fontStyle ===
                      "italic"
                        ? "active"
                        : ""
                    }
                    onClick={() =>
                      updateElement(
                        "fontStyle",
                        currentElement.fontStyle ===
                          "italic"
                          ? "normal"
                          : "italic"
                      )
                    }
                  >
                    <i>
                      I
                    </i>
                  </button>

                </div>

              </div>

              {/* UPPERCASE */}

              <div className="setting-group">

                <label>
                  Әріп форматы
                </label>

                <button
                  className={
                    currentElement?.uppercase
                      ? "uppercase-button active"
                      : "uppercase-button"
                  }
                  onClick={() =>
                    updateElement(
                      "uppercase",
                      !currentElement.uppercase
                    )
                  }
                >
                  {currentElement?.uppercase
                    ? "🔠 БАС ӘРІП — ҚОСУЛЫ"
                    : "🔡 Қалыпты әріп"}
                </button>

                <small className="hint">
                  Қосылса, Excel-дегі
                  мәтін PDF-те толық
                  БАС ӘРІППЕН шығады.
                </small>

              </div>

              {/* ALIGN */}

              <div className="setting-group">

                <label>
                  Туралау
                </label>

                <div className="align-buttons">

                  <button
                    className={
                      currentElement?.textAlign ===
                      "left"
                        ? "active"
                        : ""
                    }
                    onClick={() =>
                      updateElement(
                        "textAlign",
                        "left"
                      )
                    }
                  >
                    ≡
                  </button>

                  <button
                    className={
                      currentElement?.textAlign ===
                      "center"
                        ? "active"
                        : ""
                    }
                    onClick={() =>
                      updateElement(
                        "textAlign",
                        "center"
                      )
                    }
                  >
                    ≡
                  </button>

                  <button
                    className={
                      currentElement?.textAlign ===
                      "right"
                        ? "active"
                        : ""
                    }
                    onClick={() =>
                      updateElement(
                        "textAlign",
                        "right"
                      )
                    }
                  >
                    ≡
                  </button>

                </div>

              </div>

            </>
          )}

          {/* QR */}

          {currentElement?.id ===
            "qr" && (

            <div className="setting-group">

              <div className="label-line">

                <label>
                  QR өлшемі
                </label>

                <b>
                  {
                    currentElement?.qrSize
                  }
                  px
                </b>

              </div>

              <input
                type="range"
                min="60"
                max="300"
                value={
                  currentElement?.qrSize
                }
                onChange={(e) =>
                  updateElement(
                    "qrSize",
                    Number(
                      e.target.value
                    )
                  )
                }
              />

              <button
                className="qr-button"
                onClick={async () => {

                  if (!firstStudent) {
                    return;
                  }

                  try {

                    const qr =
                      await QRCode.toDataURL(
                        getQrText(
                          firstStudent
                        ),
                        {
                          width: 1800,
                          margin: 1,
                          errorCorrectionLevel:
                            "H",
                          color: {
                            dark: "#000000",
                            light: "#ffffff",
                          },
                        }
                      );

                    setQrPreview(qr);

                  } catch (error) {

                    console.error(
                      error
                    );

                  }

                }}
              >
                🔳 QR жаңарту
              </button>

            </div>

          )}

          {/* HELP */}

          <div className="keyboard-box">

            <b>
              ⌨️ Дәл орналастыру
            </b>

            <p>
              ← ↑ ↓ → — 1 пиксель
            </p>

            <p>
              Shift + стрелка —
              10 пиксель
            </p>

            <p>
              X / Y арқылы нақты
              координата қоюға болады.
            </p>

            <p>
              🎯 Крестик — дипломның
              нақты ортасы.
            </p>

          </div>

        </aside>

      </main>
    </div>
  );
}

export default App;