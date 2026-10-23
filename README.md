# Flappy Cen Pro

Juego arcade de navegador con Canvas y JavaScript. Es una aplicación estática: no necesita paquetes, servidor de backend ni proceso de compilación.

## Estructura

- **index.html**: página de entrada del juego.
- **src/css/game.css**: estilos y adaptación visual.
- **src/js/game.js**: física, controles, colisiones, puntuación y renderizado.
- **src/js/audio.js**: efectos de sonido y control de silencio.
- **assets/images/personaje/**: sprites del personaje.

## Ejecutar localmente

Desde la raíz del repositorio, inicia un servidor estático:

    py -m http.server 8000

Abre http://localhost:8000/ en el navegador.

No se necesita instalar paquetes ni compilar. También puedes abrir index.html directamente, aunque un servidor local reproduce mejor el comportamiento de un sitio publicado.

## Controles

- Toca o haz clic en el juego para volar.
- En teclado, usa Espacio, Flecha arriba o W.
- Pulsa M para silenciar y R para reiniciar.
- El récord se guarda en el almacenamiento local del navegador.
