function showProcessingMessage(container, timerElement, timerValueElement) {
    const startTime = Date.now();
    let dotCount = 1;
    const maxDots = 10;

    container.innerHTML = `
        <div class="processing-message">
            Идет обработка файлов<span class="dots"></span>
        </div>
    `;

    timerElement.style.display = "block";
    timerValueElement.textContent = "0.00";

    const dotsElement = container.querySelector(".dots");
    const frameRate = 500;

    const animationInterval = setInterval(() => {
        dotsElement.textContent = ".".repeat(dotCount);
        dotCount = (dotCount % maxDots) + 1;
    }, frameRate);

    const timerInterval = setInterval(() => {
        const elapsed = (Date.now() - startTime) / 1000;
        timerValueElement.textContent = elapsed.toFixed(2);
    }, 100);

    return {
        stop: () => {
            clearInterval(animationInterval);
            clearInterval(timerInterval);
            timerElement.style.display = "none";
        },
        getDuration: () => {
            return ((Date.now() - startTime) / 1000).toFixed(2);
        }
    };
}

function updateUI(container, type, options = {}) {
    const classes = {
        info: "info-message",
        success: "success-message",
        error: "error-message"
    };

    let message = "";

    if (type === "info") {
        message =
            "Пожалуйста, загрузите основной файл сводной таблицы и файл АВР, " +
            "затем нажмите кнопку 'Обработать файлы'";
    } else if (type === "success") {
        const { duration } = options;
        message = `
            Файлы успешно обработаны!<br>
            Время выполнения: ${duration} секунд.
        `;
    } else if (type === "error") {
        const { errorMessage } = options;
        message = `
            Ошибка обработки файлов: <br> 
            <div style="margin-left: 20px;">${errorMessage}<br></div>
            Пожалуйста, устраните это и попробуйте снова.
        `;
    }

    container.innerHTML = `<div class="${classes[type]}">${message}</div>`;
}

export default { showProcessingMessage, updateUI };
