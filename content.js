(function () {
    "use strict";

    // The onReady function must be run each time a new page is loaded.
    Office.onReady(function (info) {
        document.getElementById("get-data-from-selection").addEventListener("click", () => getDataFromSelection());
    });

    // Gets and displays some details about the current slide.
    async function getDataFromSelection() {
        
        try {
            const container = document.getElementById("app-body") || document.body;

            // 1. Buscamos si ya existe un iframe creado
            let iframe = document.querySelector("iframe");

            // 2. Si no existe, lo creamos dinámicamente
            if (!iframe) {
                iframe = document.createElement("iframe");
                
                // Aplicamos estilos para que ocupe el espacio del panel
                iframe.style.width = "100%";
                iframe.style.height = "500px";
                iframe.style.border = "none";
                iframe.style.marginTop = "10px";
                iframe.setAttribute("allow", "fullscreen");

                // Lo agregamos al contenedor de la interfaz
                container.appendChild(iframe);
            }

            // 3. Le asignamos el sitio web a cargar
            iframe.src = "https://es.wikipedia.org";
            await PowerPoint.run(async (context) => {
                const slides = context.presentation.getSelectedSlides();
                slides.load("items/id,items/index");
                await context.sync();

                const details = slides.items.map((slide) => ({
                    id: slide.id,
                    index: slide.index
                }));
                document.getElementById("selected-data").textContent =
                    'Hello, world! Some slide details are: ' + JSON.stringify(details);
            });
        } catch (error) {
            document.getElementById("selected-data").textContent = 'Error getting slide details.';
            console.error('Error:', error.message);
        }
    }
})();