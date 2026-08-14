function initialize() {
    const form = document.getElementById("url-form");
    const input = document.getElementById("url-input");
    const frame = document.getElementById("web-frame");
    const error = document.getElementById("error");

    const isWebUrl = (value) => {
        try {
            return /^https?:$/.test(new URL(value).protocol);
        } catch {
            return false;
        }
    };

    const savedUrl = Office.context.document.settings.get("webUrl");

    const showError = (message) => {
        error.textContent = message;
        error.hidden = false;
        form.hidden = false;
        frame.hidden = true;
    };

    const showContent = (url) => {
        form.hidden = true;
        error.hidden = true;
        frame.hidden = false;
        frame.src = url;
    };

    if (isWebUrl(savedUrl)) {
        input.value = savedUrl;
        showContent(savedUrl);
    }

    frame.addEventListener("error", () => showError("No se pudo cargar esta página. Prueba otra URL."));

    form.addEventListener("submit", (event) => {
        event.preventDefault();
        let url;

        try {
            url = new URL(input.value);
        } catch {
            showError("La URL no es válida. Prueba otra URL.");
            return;
        }

        if (!isWebUrl(url.href)) {
            showError("La URL debe comenzar con http:// o https://. Prueba otra URL.");
            return;
        }

        Office.context.document.settings.set("webUrl", url.href);
        Office.context.document.settings.saveAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                showError("No se pudo guardar la URL. Prueba otra URL.");
                return;
            }

            showContent(url.href);
        });
    });
}

Office.onReady(() => {
    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", initialize, { once: true });
    } else {
        initialize();
    }
});
