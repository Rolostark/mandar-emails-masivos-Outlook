Set WshShell = CreateObject("WScript.Shell")
' Ejecuta el servidor Node de forma silenciosa en segundo plano (sin ventana de consola)
WshShell.Run "cmd /c node server.js", 0, false
' Espera 1.5 segundos para asegurarse de que el puerto esté listo
WScript.Sleep 1500
' Abre la aplicación en el navegador por defecto
WshShell.Run "http://localhost:3000"
