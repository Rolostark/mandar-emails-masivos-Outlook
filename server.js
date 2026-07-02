const http = require('http');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const os = require('os');

const PORT = 3000;

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css',
  '.js': 'text/javascript',
  '.json': 'application/json',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.csv': 'text/csv; charset=utf-8'
};

const server = http.createServer((req, res) => {
  // Configuración de CORS para desarrollo local si es necesario
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  // Ruta API para enviar correos
  if (req.method === 'POST' && req.url === '/api/send') {
    let bodyData = '';
    req.on('data', chunk => {
      bodyData += chunk;
    });

    req.on('end', () => {
      try {
        const payload = JSON.parse(bodyData);
        handleEmailSending(payload, res);
      } catch (err) {
        res.writeHead(400, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'JSON inválido o malformado' }));
      }
    });
    return;
  }

  // Servir archivos estáticos
  if (req.method === 'GET') {
    let filePath = req.url === '/' ? '/index.html' : req.url;
    const publicDir = path.join(__dirname, 'public');
    filePath = path.join(publicDir, decodeURIComponent(filePath));

    // Validar que el archivo esté dentro del directorio público por seguridad
    if (!filePath.startsWith(publicDir)) {
      res.writeHead(403, { 'Content-Type': 'text/plain' });
      res.end('Acceso denegado');
      return;
    }

    fs.stat(filePath, (err, stats) => {
      if (err || !stats.isFile()) {
        res.writeHead(404, { 'Content-Type': 'text/plain' });
        res.end('Archivo no encontrado');
        return;
      }

      const ext = path.extname(filePath).toLowerCase();
      const contentType = MIME_TYPES[ext] || 'application/octet-stream';

      res.writeHead(200, { 'Content-Type': contentType });
      const stream = fs.createReadStream(filePath);
      stream.pipe(res);
    });
    return;
  }

  res.writeHead(405, { 'Content-Type': 'text/plain' });
  res.end('Método no permitido');
});

function handleEmailSending(payload, res) {
  const { contacts, subject, body, delay, draftMode } = payload;

  if (!contacts || !Array.isArray(contacts) || !subject || !body) {
    res.writeHead(400, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ error: 'Faltan parámetros requeridos' }));
    return;
  }

  // Establecer cabeceras para streaming de respuestas (Chunked transfer)
  res.writeHead(200, {
    'Content-Type': 'application/x-ndjson',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive',
    'Transfer-Encoding': 'chunked'
  });

  // Generar script de PowerShell adaptado con los datos recibidos
  const subjectEscaped = subject.replace(/'/g, "''");
  const bodyEscaped = body.replace(/'/g, "''");
  const delaySec = parseInt(delay) || 2;
  const isDraft = draftMode ? '$true' : '$false';
  
  // Limpiar contactos para pasarlos de forma segura en JSON
  const contactsJson = JSON.stringify(contacts, null, 2);

  // Script de PowerShell a ejecutar localmente
  const psScript = `\ufeff# SCRIPT AUTO-GENERADO POR SERVIDOR LOCAL MAILMERGE STUDIO
$outputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$subjectTemplate = @'
${subjectEscaped}
'@

$cuerpoTemplate = @'
<html>
<head>
<meta charset="utf-8">
<style>
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 14px; line-height: 1.6; color: #333333; }
</style>
</head>
<body>
${bodyEscaped}
</body>
</html>
'@

$segundosDeEspera = ${delaySec}
$modoBorrador = ${isDraft}

$contactosJson = @'
${contactsJson}
'@

$contactos = ConvertFrom-Json $contactosJson

$msg = @{ type = "log"; message = "Estableciendo canal COM con Outlook Desktop (asegúrate de tenerlo abierto)..." }
Write-Output "STATUS: $($msg | ConvertTo-Json -Compress)"

try {
    $outlook = New-Object -ComObject Outlook.Application
} catch {
    $msg = @{ type = "error"; message = "No se pudo conectar con Outlook de escritorio. Asegúrate de tener la aplicación Outlook clásica abierta y configurada en tu PC." }
    Write-Output "STATUS: $($msg | ConvertTo-Json -Compress)"
    exit 1
}

$total = $contactos.Count
$contador = 0

foreach ($c in $contactos) {
    $email = $c.__EmailDestinatario__
    if (-not $email -or $email -eq "") {
        continue
    }

    $asunto = $subjectTemplate
    $cuerpo = $cuerpoTemplate

    foreach ($prop in $c.PSObject.Properties) {
        $key = $prop.Name
        if ($key -eq "__EmailDestinatario__") { continue }
        $val = if ($prop.Value -eq $null) { "" } else { $prop.Value.ToString() }
        $asunto = $asunto.Replace("{{$key}}", $val)
        $cuerpo = $cuerpo.Replace("{{$key}}", $val)
    }

    try {
        $mail = $outlook.CreateItem(0)
        $mail.To = $email
        $mail.Subject = $asunto
        $mail.HTMLBody = $cuerpo
        
        if ($modoBorrador) {
            $mail.Save()
            $statusMsg = "Borrador guardado"
        } else {
            $mail.Send()
            $statusMsg = "Correo enviado"
        }
        $contador++
        $msg = @{
            type = "progress"
            index = $contador
            total = $total
            email = $email
            success = $true
            message = $statusMsg
        }
        Write-Output "STATUS: $($msg | ConvertTo-Json -Compress)"
    } catch {
        $err = $_.Exception.Message
        $msg = @{
            type = "progress"
            index = ($contador + 1)
            total = $total
            email = $email
            success = $false
            message = $err
        }
        Write-Output "STATUS: $($msg | ConvertTo-Json -Compress)"
    }

    if ($contador -lt $total) {
        Start-Sleep -Seconds $segundosDeEspera
    }
}

$msg = @{ type = "done"; total = $contador }
Write-Output "STATUS: $($msg | ConvertTo-Json -Compress)"
`;

  // Crear archivo temporal
  const tempFileName = `.temp_run_${Date.now()}.ps1`;
  const tempFilePath = path.join(os.tmpdir(), tempFileName);

  fs.writeFile(tempFilePath, psScript, { encoding: 'utf8' }, (err) => {
    if (err) {
      res.write(JSON.stringify({ type: 'error', message: 'No se pudo crear el archivo temporal de ejecución.' }) + '\n');
      res.end();
      return;
    }

    // Ejecutar PowerShell con la política bypass en modo no interactivo
    const child = spawn('powershell.exe', [
      '-NoProfile',
      '-ExecutionPolicy', 'Bypass',
      '-File', tempFilePath
    ]);

    let buffer = '';

    child.stdout.on('data', (data) => {
      buffer += data.toString('utf8');
      let lines = buffer.split('\n');
      // Mantener la última línea incompleta en el buffer
      buffer = lines.pop();

      for (let line of lines) {
        line = line.trim();
        if (line.startsWith('STATUS: ')) {
          const jsonStr = line.substring(8).trim();
          try {
            const parsed = JSON.parse(jsonStr);
            res.write(JSON.stringify(parsed) + '\n');
          } catch (e) {
            res.write(JSON.stringify({ type: 'log', message: line }) + '\n');
          }
        } else if (line !== '') {
          res.write(JSON.stringify({ type: 'log', message: line }) + '\n');
        }
      }
    });

    child.stderr.on('data', (data) => {
      res.write(JSON.stringify({ type: 'error', message: data.toString('utf8') }) + '\n');
    });

    child.on('error', (spawnError) => {
      res.write(JSON.stringify({ type: 'error', message: `No se pudo iniciar el proceso de PowerShell: ${spawnError.message}` }) + '\n');
    });

    child.on('close', (code) => {
      // Limpiar archivo temporal al terminar
      fs.unlink(tempFilePath, () => {});
      res.end();
    });
  });
}

server.listen(PORT, () => {
  console.log(`Servidor local corriendo en http://localhost:${PORT}`);
  // Abrir automáticamente el navegador por defecto en Windows
  const { exec } = require('child_process');
  exec('start http://localhost:' + PORT);
});

