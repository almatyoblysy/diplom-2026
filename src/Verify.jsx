import { useEffect, useState } from "react";

function Verify() {
  const [student, setStudent] = useState(null);
  const [error, setError] = useState("");

  useEffect(() => {
    try {
      // URL:
      // /diplom/verify/0001

      const path = window.location.pathname;
      const parts = path.split("/").filter(Boolean);

      const verifyIndex = parts.indexOf("verify");

      if (verifyIndex === -1 || !parts[verifyIndex + 1]) {
        setError("Диплом нөмірі көрсетілмеген.");
        return;
      }

      const registration = decodeURIComponent(
        parts[verifyIndex + 1]
      );

      // Уақытша тест мәліметі.
      // Кейін мұнда Excel-ден жасалған JSON қосамыз.
      const demoStudent = {
        "Аудан": "Қарасай ауданы",
        "Мекеме атауы": "Абай атындағы мектеп-гимназиясы",
        "Жетекшісінің аты-жөні": "Тест Жетекші",
        "Оқушының аты-жөні": "Тест Оқушы",
        "Тіркеу №": registration,
        "Байқау атауы": "Тест байқауы",
        "Пәні": "Информатика",
        "Номинация": "Робототехника",
        "Түрі": "Диплом",
        "Жүлделі орын": "I орын",
        "Өткізу бұйрық номері/күні":
          "№01 / 01.01.2026",
      };

      setStudent(demoStudent);
    } catch (err) {
      console.error(err);
      setError("Диплом мәліметін оқу кезінде қате болды.");
    }
  }, []);

  if (error) {
    return (
      <div className="verify-page">
        <div className="verify-card">
          <div className="verify-icon">❌</div>

          <h1>Диплом табылмады</h1>

          <p>{error}</p>
        </div>
      </div>
    );
  }

  if (!student) {
    return (
      <div className="verify-page">
        <div className="verify-card">
          <div className="verify-icon">⏳</div>

          <h1>Тексерілуде...</h1>
        </div>
      </div>
    );
  }

  return (
    <div className="verify-page">
      <div className="verify-card">
        <div className="verify-success">
          ✓
        </div>

        <h1>Диплом расталды</h1>

        <p className="verify-subtitle">
          «Жас Дарын» дипломын тексеру
        </p>

        <div className="verify-info">
          <div className="verify-row">
            <span>Оқушының аты-жөні</span>
            <strong>
              {student["Оқушының аты-жөні"]}
            </strong>
          </div>

          <div className="verify-row">
            <span>Тіркеу №</span>
            <strong>
              {student["Тіркеу №"]}
            </strong>
          </div>

          <div className="verify-row">
            <span>Аудан</span>
            <strong>
              {student["Аудан"]}
            </strong>
          </div>

          <div className="verify-row">
            <span>Мекеме атауы</span>
            <strong>
              {student["Мекеме атауы"]}
            </strong>
          </div>

          <div className="verify-row">
            <span>Жетекшісінің аты-жөні</span>
            <strong>
              {student["Жетекшісінің аты-жөні"]}
            </strong>
          </div>

          <div className="verify-row">
            <span>Байқау атауы</span>
            <strong>
              {student["Байқау атауы"]}
            </strong>
          </div>

          <div className="verify-row">
            <span>Пәні</span>
            <strong>
              {student["Пәні"]}
            </strong>
          </div>

          <div className="verify-row">
            <span>Номинация</span>
            <strong>
              {student["Номинация"]}
            </strong>
          </div>

          <div className="verify-row">
            <span>Түрі</span>
            <strong>
              {student["Түрі"]}
            </strong>
          </div>

          <div className="verify-row">
            <span>Жүлделі орын</span>
            <strong>
              {student["Жүлделі орын"]}
            </strong>
          </div>

          <div className="verify-row">
            <span>
              Өткізу бұйрық номері/күні
            </span>
            <strong>
              {student["Өткізу бұйрық номері/күні"]}
            </strong>
          </div>
        </div>

        <div className="verify-footer">
          🔐 Бұл мәлімет QR-код арқылы
          дипломды тексеру үшін көрсетілді.
        </div>
      </div>
    </div>
  );
}

export default Verify;