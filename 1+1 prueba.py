import subprocess

# El resto de tu código de Python va aquí
resultado = 1 + 1
print(resultado)

# Al final de tu script, añade lo siguiente para ejecutar el atajo en un proceso separado
subprocess.Popen(["/usr/bin/shortcuts", "run", "NotificacionPython"])
