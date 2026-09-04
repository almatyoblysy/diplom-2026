import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import QRCode from "qrcode";
import jsPDF from "jspdf";

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
      visible: false,
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
      y: 535,
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
      visible: false,
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
      qrSize: 120,
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
  const [overlayPreview, setOverlayPreview] = useState(false);
  const [overlayPdfLoading, setOverlayPdfLoading] = useState(false);

  // =========================================================
  // PREVIEW
  // =========================================================

  const [qrPreview, setQrPreview] = useState("");

  const editorRef = useRef(null);

  // =========================================================
  // FONTS — keep the editor/PDF fonts loaded before rendering
  // =========================================================

  useEffect(() => {
    const fontUrl =
      "https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;600;700&family=Inter:wght@400;600;700&family=Lora:wght@400;600;700&family=Merriweather:wght@400;700&family=Montserrat:wght@400;600;700&family=Nunito:wght@400;600;700&family=Open+Sans:wght@400;600;700&family=Oswald:wght@400;500;600;700&family=Playfair+Display:wght@400;600;700&family=Raleway:wght@400;600;700&display=swap";

    if (!document.querySelector('link[data-diploma-fonts="true"]')) {
      const link = document.createElement("link");
      link.rel = "stylesheet";
      link.href = fontUrl;
      link.dataset.diplomaFonts = "true";
      document.head.appendChild(link);
    }
  }, []);

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

  const columnAliases = {
    district: ["Аудан", "Аудан/қала", "Қала/аудан", "Аудан немесе қала"],
    institution: ["Мекеме атауы", "Мектеп атауы", "Мектеп", "Білім беру ұйымы", "Білім беру мекемесі"],
    leader: ["Жетекшісінің аты-жөні", "Жетекшінің аты-жөні", "Жетекші", "Жетекшісі"],
    name: ["Оқушының аты-жөні", "Оқушы", "Аты-жөні", "ФИО"],
    registration: ["Тіркеу №", "Тіркеу нөмірі", "Тіркеу номері", "Регистрационный №", "Рег. №"],
    competition: ["Байқау атауы", "Байқау", "Конкурс атауы", "Конкурс"],
    subject: ["Пәні", "Пән", "Предмет"],
    nomination: ["Номинация", "Номинациясы"],
    type: ["Түрі", "Тип", "Типі"],
    place: ["Жүлделі орын", "Орын", "Место"],
    order: ["Өткізу бұйрық номері/күні", "Бұйрық номері/күні", "Бұйрық", "Приказ номері/күні", "Приказ"],
  };

  const normalizeHeader = (value) => {
    return String(value ?? "")
      .toLowerCase()
      .replace(/№/g, "no")
      .replace(/[«»\"'`]/g, "")
      .replace(/[./\\_-]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  };

  const cleanStudentValue = (value) => {
    if (value === null || value === undefined) return "";
    return String(value).replace(/\s+/g, " ").trim();
  };

  const findColumn = (student, wanted, aliases = []) => {
    if (!student) return "";
    const keys = Object.keys(student);
    for (const name of [wanted, ...aliases].filter(Boolean)) {
      const normalizedWanted = normalizeHeader(name);
      const exact = keys.find((key) => normalizeHeader(key) === normalizedWanted);
      if (exact) return cleanStudentValue(student[exact]);
    }
    return "";
  };

  const getStudentValue = (student, id) => {
    if (!student) return "";
    return findColumn(student, columnMap[id], columnAliases[id] || []);
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

      setErrorMessage(
        "Дизайн жүктеу кезінде қате болды."
      );
    }
  };

  // =========================================================
  // STUDENT PREVIEW
  // =========================================================

  const firstStudent = students.length
    ? students[0]
    : null;

  // =========================================================
  // TEXT
  // =========================================================

  const getElementText = (element, student) => {
    if (!student || element.id === "qr") {
      return element.id === "qr" ? "QR" : "";
    }

    let text = String(
      getStudentValue(student, element.id) ?? ""
    );

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

    const qrFields = [
      ["Аудан", getStudentValue(student, "district")],
      ["Мекеме атауы", getStudentValue(student, "institution")],
      ["Жетекшісінің аты-жөні", getStudentValue(student, "leader")],
      ["Оқушының аты-жөні", getStudentValue(student, "name")],
      ["Тіркеу №", getStudentValue(student, "registration")],
      ["Байқау атауы", getStudentValue(student, "competition")],
      ["Пәні", getStudentValue(student, "subject")],
      ["Номинация", getStudentValue(student, "nomination")],
      ["Түрі", getStudentValue(student, "type")],
      ["Жүлделі орын", getStudentValue(student, "place")],
      ["Өткізу бұйрық номері/күні", getStudentValue(student, "order")],
    ];

    return qrFields
      .filter(([, value]) => String(value ?? "").trim() !== "")
      .map(([label, value]) => `${label}: ${String(value).trim()}`)
      .join("\n");
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

    const rect =
      editorRef.current.getBoundingClientRect();

    const scaleX =
      canvasSize.width / rect.width;

    const scaleY =
      canvasSize.height / rect.height;

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

    const rect =
      editorRef.current.getBoundingClientRect();

    const scaleX =
      canvasSize.width / rect.width;

    const scaleY =
      canvasSize.height / rect.height;

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
    // 3x render: PDF/print quality is much better.
    const SCALE = 3;

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

    // White base
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    // Diploma background
    ctx.drawImage(
      backgroundImage,
      0,
      0,
      canvas.width,
      canvas.height
    );

    // =======================================================
    // QR — large source image, same physical size
    // =======================================================

    let qrImage = null;

    const qrElement = elements.find(
      (element) => element.id === "qr"
    );

    if (qrElement && qrElement.visible !== false) {
      const qrText = getQrText(student);

      if (qrText) {
        const qrDataUrl = await QRCode.toDataURL(
          qrText,
          {
            // Large source makes the small QR modules crisp in print.
            width: 1800,
            margin: 1,
            errorCorrectionLevel: "H",
            color: {
              dark: "#000000",
              light: "#ffffff",
            },
          }
        );

        qrImage = await loadQrImage(qrDataUrl);
      }
    }

    // =======================================================
    // ELEMENTS
    // =======================================================

    for (const element of elements) {
      if (element.visible === false) continue;

      // ---------------- QR ----------------
      if (element.id === "qr") {
        if (!qrImage) continue;

        const size =
          (Number(element.qrSize) || 120) * SCALE;

        const x = element.x * SCALE;
        const y = element.y * SCALE;

        // QR modules should stay sharp.
        ctx.save();
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

      // ---------------- TEXT ----------------
      const text = getElementText(element, student);

      if (!text) continue;

      ctx.save();

      const weight =
        element.fontWeight === "bold" ? "bold" : "normal";

      const style =
        element.fontStyle === "italic" ? "italic" : "normal";

      ctx.font =
        `${style} ${weight} ${element.fontSize * SCALE}px "${element.fontFamily}"`;

      ctx.fillStyle = element.color;
      ctx.textAlign = element.textAlign || "center";
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
      if (!students.length) {
        throw new Error(
          "Алдымен Excel файлын жүктеңіз."
        );
      }

      if (!designUrl) {
        throw new Error(
          "Алдымен диплом дизайнын жүктеңіз."
        );
      }

      setPdfLoading(true);

      if (document.fonts?.ready) {
        await document.fonts.ready;
      }

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
          JPEG қолданамыз:
          500+ диплом кезінде PDF көлемі тым үлкен болып
          кетпеуі үшін.
        */
        const imageData =
          diplomaCanvas.toDataURL(
            "image/png"
          );

        if (i > 0) {
          pdf.addPage(
            "a4",
            isLandscape
              ? "landscape"
              : "portrait"
          );
        }

        pdf.addImage(
          imageData,
          "PNG",
          0,
          0,
          pageWidth,
          pageHeight,
          undefined,
          "NONE"
        );

        /*
          Canvas-ты босатамыз.
        */
        diplomaCanvas.width = 1;
        diplomaCanvas.height = 1;

        /*
          Браузерге демалу үшін әр 10 беттен кейін
          кішкене кідіріс.
        */
        if (
          i % 10 === 0 &&
          i !== 0
        ) {
          await new Promise(
            (resolve) =>
              setTimeout(resolve, 20)
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
  // ДИПЛОМ ҮСТІНЕ — OVERLAY MODE
  // =========================================================

  // Диплом үстіне басылатын барлық өзгермелі деректер.
  // Түрі мен жүлделі орын бұрынғы логика бойынша фондық дизайнның
  // өзінде қалуы үшін overlay-ге кірмейді.
  const isOverlayElement = (element) => {
    if (!element) return false;
    if (element.visible === false) return false;
    if (element.id === "type") return false;
    if (element.id === "place") return false;
    return true;
  };

  const drawOverlayElements = async (ctx, student, scale = 1) => {
    let qrImage = null;

    const qrElement = elements.find((element) => element.id === "qr");

    if (qrElement && qrElement.visible !== false) {
      const qrText = getQrText(student);

      if (qrText) {
        const qrDataUrl = await QRCode.toDataURL(qrText, {
          width: 1800,
          margin: 1,
          errorCorrectionLevel: "H",
          color: {
            dark: "#000000",
            light: "#ffffff",
          },
        });

        qrImage = await loadQrImage(qrDataUrl);
      }
    }

    for (const element of elements) {
      if (!isOverlayElement(element)) continue;

      if (element.id === "qr") {
        if (!qrImage) continue;

        const size = (Number(element.qrSize) || 120) * scale;
        const x = element.x * scale;
        const y = element.y * scale;

        ctx.save();
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

      const text = getElementText(element, student);
      if (!text) continue;

      ctx.save();

      const weight = element.fontWeight === "bold" ? "bold" : "normal";
      const style = element.fontStyle === "italic" ? "italic" : "normal";

      ctx.font = `${style} ${weight} ${element.fontSize * scale}px "${element.fontFamily}"`;
      ctx.fillStyle = element.color;
      ctx.textAlign = element.textAlign || "center";
      ctx.textBaseline = "middle";

      ctx.fillText(
        text,
        element.x * scale,
        element.y * scale
      );

      ctx.restore();
    }
  };

  const createOverlayCanvas = async (student) => {
    const SCALE = 3;

    const canvas = document.createElement("canvas");
    canvas.width = canvasSize.width * SCALE;
    canvas.height = canvasSize.height * SCALE;

    const ctx = canvas.getContext("2d", {
      alpha: false,
      willReadFrequently: false,
    });

    if (!ctx) {
      throw new Error("Overlay canvas құрылғысын жасау мүмкін болмады.");
    }

    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = "high";

    // ТЕК АҚ ФОН. Диплом дизайны бұл PDF-ке кірмейді.
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    await drawOverlayElements(ctx, student, SCALE);

    return canvas;
  };

  const generateOverlayPDF = async () => {
    if (overlayPdfLoading || pdfLoading) return;

    setErrorMessage("");
    setPdfProgress("");

    try {
      if (!students.length) {
        throw new Error("Алдымен Excel файлын жүктеңіз.");
      }

      setOverlayPdfLoading(true);

      if (document.fonts?.ready) {
        await document.fonts.ready;
      }

      const isLandscape = orientation === "landscape";
      const pdf = new jsPDF({
        orientation: isLandscape ? "landscape" : "portrait",
        unit: "mm",
        format: "a4",
        compress: true,
      });

      const pageWidth = isLandscape ? 297 : 210;
      const pageHeight = isLandscape ? 210 : 297;

      for (let i = 0; i < students.length; i++) {
        const student = students[i];

        setPdfProgress(
          `🖨️ Диплом үстіне арналған дерек: ${i + 1} / ${students.length}`
        );

        const overlayCanvas = await createOverlayCanvas(student);
        const imageData = overlayCanvas.toDataURL("image/png");

        if (i > 0) {
          pdf.addPage(
            "a4",
            isLandscape ? "landscape" : "portrait"
          );
        }

        pdf.addImage(
          imageData,
          "PNG",
          0,
          0,
          pageWidth,
          pageHeight,
          undefined,
          "NONE"
        );

        overlayCanvas.width = 1;
        overlayCanvas.height = 1;

        if (i % 10 === 0 && i !== 0) {
          await new Promise((resolve) => setTimeout(resolve, 20));
        }
      }

      pdf.save(`Диплом_үстіне_${students.length}_бет.pdf`);

      setPdfProgress(
        `✅ ДАЙЫН! ${students.length} ақ overlay беті жасалды.`
      );
    } catch (error) {
      console.error("OVERLAY PDF ERROR:", error);

      setErrorMessage(
        "❌ Диплом үстіне PDF жасау кезінде қате болды: " +
          (error?.message || String(error))
      );
      setPdfProgress("");
    } finally {
      setOverlayPdfLoading(false);
    }
  };

  const renderOverlayPreviewElement = (element) => {
    const text = getElementText(element, firstStudent);

    if (element.id === "qr") {
      return qrPreview ? (
        <img
          src={qrPreview}
          alt="QR"
          draggable="false"
          className="overlay-preview-qr"
        />
      ) : (
        <div className="overlay-qr-empty">QR</div>
      );
    }

    return text || element.label;
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

    if (element.id === "qr") {
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

    return text || element.label;
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
                            (selected
                              ? "selected-element"
                              : "") +
                            (element.id ===
                            "qr"
                              ? " qr-element"
                              : "")
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

                <div className="pdf-buttons">
                  <button
                    className="pdf-button"
                    onClick={generateAllPDF}
                    disabled={pdfLoading || overlayPdfLoading}
                  >
                    {pdfLoading
                      ? "⏳ Жасалып жатыр..."
                      : "📥 БАРЛЫҚ ДИПЛОМДЫ PDF ЖАСАУ"}
                  </button>

                  <button
                    className="overlay-preview-button"
                    onClick={() => setOverlayPreview(true)}
                    disabled={pdfLoading || overlayPdfLoading}
                  >
                    🖨️ Диплом үстіне көру
                  </button>

                  <button
                    className="overlay-pdf-button"
                    onClick={generateOverlayPDF}
                    disabled={pdfLoading || overlayPdfLoading}
                  >
                    {overlayPdfLoading
                      ? "⏳ Дайындалып жатыр..."
                      : "📄 ТЕК ДЕРЕКТЕР PDF — ДИПЛОМ ҮСТІНЕ"}
                  </button>
                </div>

                {pdfProgress && (
                  <div className="pdf-progress">
                    {pdfProgress}
                  </div>
                )}
              </div>
            )}

          {/* =================================================
              ДИПЛОМ ҮСТІНЕ PREVIEW
          ================================================= */}
          {overlayPreview && students.length > 0 && designUrl && (
            <div className="overlay-card">
              <div className="overlay-header">
                <div>
                  <h2>🖨️ Диплом үстіне</h2>
                  <p>
                    Ақ қабатта тек өзгермелі деректер көрсетіледі. Бұл қабат
                    қол қойылған бос дипломның үстінен басып шығаруға арналған.
                  </p>
                </div>
                <button
                  className="overlay-close-button"
                  onClick={() => setOverlayPreview(false)}
                >
                  ✕ Жабу
                </button>
              </div>

              <div className="overlay-info">
                <b>Алдымен:</b> бос дипломды басып шығарып, қол қойып алыңыз.
                <br />
                <b>Содан кейін:</b> осы ақ overlay PDF-ін дәл сол дипломның
                үстінен басып шығарасыз.
              </div>

              <div className="overlay-stage-wrap">
                <div
                  className="overlay-stage"
                  style={{
                    width: canvasSize.width,
                    height: canvasSize.height,
                    aspectRatio: `${canvasSize.width}/${canvasSize.height}`,
                  }}
                >
                  <img
                    src={designUrl}
                    alt="Диплом фоны"
                    className="overlay-background"
                    draggable="false"
                  />

                  <div className="overlay-tint" />

                  {elements.map((element) => {
                    if (!isOverlayElement(element)) return null;

                    const textAlign = element.textAlign || "center";
                    let transform = "translate(-50%, -50%)";

                    if (textAlign === "left") {
                      transform = "translate(0, -50%)";
                    } else if (textAlign === "right") {
                      transform = "translate(-100%, -50%)";
                    }

                    return (
                      <div
                        key={element.id}
                        className={
                          "overlay-element" +
                          (element.id === "qr" ? " overlay-qr-element" : "")
                        }
                        style={{
                          left: element.x,
                          top: element.y,
                          transform,
                          fontFamily: element.fontFamily,
                          fontSize: element.fontSize,
                          fontWeight: element.fontWeight,
                          fontStyle: element.fontStyle,
                          color: element.color,
                          textAlign,
                          width: element.id === "qr" ? element.qrSize : "max-content",
                          height: element.id === "qr" ? element.qrSize : "auto",
                        }}
                      >
                        {renderOverlayPreviewElement(element)}
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className="overlay-actions">
                <button
                  className="overlay-preview-button"
                  onClick={() => setOverlayPreview(false)}
                >
                  🎨 Диплом макетіне қайту
                </button>

                <button
                  className="overlay-pdf-button"
                  onClick={generateOverlayPDF}
                  disabled={overlayPdfLoading || pdfLoading}
                >
                  {overlayPdfLoading
                    ? "⏳ Overlay PDF жасалып жатыр..."
                    : "📄 ТЕК ДЕРЕКТЕР PDF — ДИПЛОМ ҮСТІНЕ"}
                </button>
              </div>
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
            {elements
              .filter(
                (element) =>
                  element.id !== "type" &&
                  element.id !== "place"
              )
              .map(
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
                    {element.label}
                  </button>
                )
              )}
          </div>

          <div className="selected-name">
            {currentElement?.label}
          </div>

          {/* X Y */}

          <div className="coordinates">
            <div className="coordinate">
              <span>X</span>

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
              <span>Y</span>

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

                  <option value="Inter">Inter</option>
                  <option value="Montserrat">Montserrat</option>
                  <option value="Playfair Display">Playfair Display</option>
                  <option value="Cormorant Garamond">Cormorant Garamond</option>
                  <option value="Lora">Lora</option>
                  <option value="Merriweather">Merriweather</option>
                  <option value="Nunito">Nunito</option>
                  <option value="Open Sans">Open Sans</option>
                  <option value="Oswald">Oswald</option>
                  <option value="Raleway">Raleway</option>
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
                    <b>B</b>
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
                    <i>I</i>
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
                  if (!firstStudent)
                    return;

                  try {
                    const qr =
                      await QRCode.toDataURL(
                        getQrText(
                          firstStudent
                        ),
                        {
                          width: 500,
                          margin: 1,
                          errorCorrectionLevel:
                            "H",
                        }
                      );

                    setQrPreview(qr);
                  } catch (
                    error
                  ) {
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