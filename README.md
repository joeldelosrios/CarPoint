es un sistema web desarrollado con PHP y MySQL que cubre una serie de requerimientos básicos
para realizar una factura comercial a nuestros clientes que va desde el manejo de usuarios, creación de productos, clientes y facturas 
dentro del mismo sistema.

Instalación en windows (servidor local) 

1-Descargar los archivos fuentes del sistema

2- Copiar y descomprimir el archivo en la carpeta c:\xampp\htdocs, al final tendras una carpeta llamada “Carpoint”, a la cual podras acceder desde el navegador como: http://localhost/Carpoint/

3- Crear una base de datos usando PHPMyAdmin accediendo a la url siguiente: http://localhost/phpmyadmin/

4- Importar las tablas de la base de datos para ello vamos a buscar el archivo "Carpoint.sql" en el directorio root de nuestro sistema, una vez localizado procedemos a hacer la importacion de los datos desde PHPMyAdmin

5- Configurar los datos de conexión a la base de datos editando el archivo de configuración que se encuentra en la siguiente ruta: Carpoint/config/db.php

6- Vista web: http://localhost/Carpoint/

7- Datos de acceso por defecto: usuario: admin y contraseña: admin

