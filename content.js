document.addEventListener("DOMContentLoaded", () => {
    const form = document.getElementById("url-form");
    const input = document.getElementById("url-input");
    const frame = document.getElementById("web-frame");
    const error = document.getElementById("error");

    form.addEventListener("submit", (event) => {
        event.preventDefault();
        let url;

        try {
            url = new URL(input.value);
        } catch {
            error.textContent = "Escribe una URL válida.";
            return;
        }

        if (!/^https?:$/.test(url.protocol)) {
            error.textContent = "La URL debe comenzar con http:// o https://.";
            return;
        }

        error.textContent = "";
        frame.src = url.href;
    });
});
