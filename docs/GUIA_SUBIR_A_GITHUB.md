# Guía: Cómo subir este proyecto a GitHub

## 1. Crear el repositorio en GitHub

1. Ve a [github.com](https://github.com) e inicia sesión.
2. Haz clic en **"New repository"** (botón verde).
3. Configura:
   - **Repository name:** `CIAC-SistemaFacturacion`
   - **Description:** `Sistema de gestión de facturas administrativas desarrollado en Excel VBA para el CIAC`
   - **Visibility:** Public (para portafolio) o Private
   - **NO** marques "Add a README file" (ya tienes uno)
4. Haz clic en **"Create repository"**.

---

## 2. Subir el proyecto desde tu computadora

### Opción A — Con Git (línea de comandos)

```bash
# En la carpeta del proyecto
git init
git add .
git commit -m "feat: versión inicial del sistema de facturación CIAC"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/CIAC-SistemaFacturacion.git
git push -u origin main
```

### Opción B — Subida directa desde GitHub (sin instalar Git)

1. En la página del repositorio recién creado, haz clic en **"uploading an existing file"**.
2. Arrastra todos los archivos y carpetas del proyecto.
3. Escribe un mensaje de commit: `"Versión inicial del sistema"`.
4. Haz clic en **"Commit changes"**.

> **Nota:** GitHub no permite subir carpetas vacías. Asegúrate de que todas las carpetas tengan al menos un archivo.

---

## 3. Recomendaciones para el portafolio

### Agrega Topics (etiquetas) al repositorio
Ve a la página del repositorio → ícono de engranaje junto a "About" → agrega:
```
excel  vba  excel-vba  automation  billing-system  dashboard  macro
```

### Activa GitHub Pages para el README (opcional)
No es necesario para este tipo de proyecto, pero si quieres una vista web del README ya está formateado para lucir bien.

### Pin the repository
En tu perfil de GitHub, haz clic en **"Customize your pins"** y agrega este repositorio para que aparezca destacado.

---

## 4. Estructura de commits recomendada

Si en el futuro actualizas el proyecto, usa mensajes de commit descriptivos:

```
feat: agregar nueva funcionalidad
fix: corregir error en validación de cédula
docs: actualizar manual de usuario
refactor: optimizar función de sincronización
style: mejorar formato del dashboard
```
