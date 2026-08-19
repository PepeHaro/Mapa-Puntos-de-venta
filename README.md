# Mapa Puntos de venta

Mapa interactivo de los puntos de venta de distribuidores y marcas de acabados
para construcción en México y Latinoamérica.

**Ver el mapa:** https://pepeharo.github.io/Mapa-Puntos-de-venta/

## Qué contiene

`index.html` es un archivo autocontenido: trae dentro los datos, la geometría de
los estados y todo el código. No carga nada de internet, así que funciona igual
publicado en la web, abierto con doble clic o sin conexión.

## Cómo se arma

Los datos se mantienen en hojas de cálculo internas y se compilan dentro de
`index.html` con un script que no forma parte de este repositorio. El bloque de
datos se reemplaza completo en cada actualización; el resto del archivo —
estilos, filtros, geometría — no se toca.

El compilador también señala las sucursales cuyas coordenadas caen fuera del
estado declarado, con un margen de 20 km para no marcar las zonas conurbadas.

## Fuentes

Base de datos de puntos de venta de distribuidores y marcas de acabados para
construcción.

## Créditos

Divisiones estatales de México: [amCharts geodata](https://www.amcharts.com/),
usadas bajo su licencia linkware. El texto de la licencia va incluido como
comentario dentro de `index.html`.
