# Despliegue web de Puerto A&M | Business Intelligence

Recomendación para compartirlo con una clase: **Render (plan gratis)**. Da una URL HTTPS
pública y estable, corre FastAPI sin Docker y se conecta directo a un repo de GitHub.

> **Nota sobre GitHub Pages:** no sirve para este proyecto. Pages solo publica HTML/CSS/JS
> estático; aquí el backend Python calcula los KPIs desde el Excel y responde el chat.
> El repo ya existe en https://github.com/psiquipitirija/puerto-am-bi y mientras tanto
> hay una demo pública temporal en
> https://lamps-contacted-refine-camel.trycloudflare.com (depende de que la
> PC y el túnel sigan encendidos).

## Opción recomendada: Render + GitHub

### 1. Crea el repositorio en GitHub
1. Entra a https://github.com/new.
2. Nombre sugerido: `puerto-am-bi`.
3. **No** marques “Add a README”, `.gitignore` ni licencia (ya los tenemos aquí).
4. Crea el repositorio y copia su URL (`https://github.com/TU_USUARIO/puerto-am-bi.git`).

Si no tienes cuenta, crea una gratuita; el plan gratis de Render también es gratuito.

### 2. Sube el código
Desde esta carpeta (`puerto-am-bi`):

```powershell
git init
git add .
git commit -m "Puerto A&M BI listo para despliegue"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/puerto-am-bi.git
git push -u origin main
```

La primera vez, Git puede abrir el navegador o pedir tus credenciales de GitHub.
Si ya hiciste `git init` y el commit desde esta sesión, solo necesitas las dos últimas líneas.

### 3. Crea el servicio en Render
1. Entra a https://dashboard.render.com/ y haz login con GitHub.
   Atajo: abre https://render.com/deploy?repo=https://github.com/psiquipitirija/puerto-am-bi
   (el botón «Deploy to Render» del README hace lo mismo).
2. **New → Web Service** → selecciona `puerto-am-bi` → **Connect**.
3. Render detecta automáticamente `render.yaml`:
   - Runtime: Python 3.12.7
   - Build: `pip install -r requirements.txt`
   - Start: `uvicorn app.main:app --host 0.0.0.0 --port $PORT`
   - Health check: `/api/health`
4. **Create Web Service**. En 2–4 minutos tendrás:
   `https://puerto-am-bi.onrender.com`

### 4. Verificación
- `https://puerto-am-bi.onrender.com/api/health` → `{"status":"ok","excel_exists":true,...}`
- El dashboard debe mostrar los 8 KPIs y las 6 gráficas.
- Comparte la URL del dashboard con la clase.

## Importante antes de compartir

- **Plan gratis:** el servicio se duerme tras ~15 minutos sin visitas. La primera carga
  después de eso puede tardar 30–60 s. Avisa a la clase que espere un poco.
- **Sin autenticación:** cualquiera con la URL puede ver el panel. Los datos son una
  simulación académica con cifras ficticias, así que no hay información confidencial.
- **El Excel en la nube es una copia:** `data/Puerto_AM_Base_Datos_DEMO.xlsx` se sube al
  repositorio. Si editas el Excel en tu PC, debes hacer `git add data/Puerto_AM_Base_Datos_DEMO.xlsx`,
  `git commit` y `git push` para que el despliegue muestre los cambios. La nube no puede
  escribir en tu archivo local.
- **Chat:** por defecto usa el motor local (sin API key). Si quieres el modelo externo,
  agrega `PUERTOAM_LLM_API_KEY` y, opcionalmente, `PUERTOAM_LLM_MODEL` en Environment
  del servicio en Render.

## Alternativa rápida: túnel temporal (sin cuenta)

Si solo necesitas el link para una presentación hoy y no quieres crear cuentas:

```powershell
cloudflared tunnel --url http://127.0.0.1:8124
```

Da una URL tipo `https://xxxx-xxxx.trycloudflare.com`. Es pública, temporal y deja de
funcionar al cerrar el túnel o apagar la PC. Requiere que este dashboard siga corriendo
en `127.0.0.1:8124`. Para mayor durabilidad, usa Render.
