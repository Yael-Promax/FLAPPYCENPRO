# Flappy Cen Pro

Juego arcade de navegador inspirado en Flappy Bird. El juego está hecho con Canvas y JavaScript sin dependencias ni proceso de compilación.

## Estructura

- **proyecto/index.html**: entrada del juego.
- **proyecto/src/css/game.css**: estilos y adaptación visual.
- **proyecto/src/js/game.js**: física, controles, colisiones, puntuación y renderizado.
- **proyecto/src/js/audio.js**: efectos de sonido y control de silencio.
- **proyecto/assets/images/personaje/**: sprites del personaje.
- **presentacion/**: presentación del ecosistema, su contenido JSON y el script PowerShell que genera el archivo PPTX.

## Abrir el juego

Desde la carpeta raíz del repositorio, inicia un servidor estático:

    py -m http.server 8000

Después abre:

    http://localhost:8000/proyecto/

No se necesita instalar paquetes ni compilar el juego. También puedes abrir proyecto/index.html directamente, aunque un servidor local reproduce mejor el comportamiento de un sitio publicado.

## Controles

- Toca o haz clic en el juego para volar.
- En teclado, usa Espacio, Flecha arriba o W.
- Pulsa M para silenciar y R para reiniciar.
- El récord se guarda en el almacenamiento local del navegador.

## Presentación

- Vista HTML: **presentacion/presentacion-ecosistema.html**.
- Contenido editable: **presentacion/contenido-pptx.json**.
- Generador: **presentacion/build-pptx.ps1**. El archivo PPTX se escribe en la misma carpeta.
