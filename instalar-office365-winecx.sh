#!/bin/bash
set -e
set -o pipefail

echo "==============================================="
echo "  Instalador automático Office 365 - WineCX"
echo "==============================================="

# ---------------------------------------------------------
# 1) Ir a carpeta Descargas
# ---------------------------------------------------------
cd "$HOME/Descargas"

# ---------------------------------------------------------
# 2) Descomprimir archivo MSO365.zip
# ---------------------------------------------------------
unzip -o MSO365.zip

# ---------------------------------------------------------
# 3) Entrar al contenido descomprimido
# ---------------------------------------------------------
cd "$HOME/Descargas/MSO365"

# ---------------------------------------------------------
# 4) Habilitar arquitectura de 32 bits
# ---------------------------------------------------------
sudo dpkg --add-architecture i386
sudo apt update

# ---------------------------------------------------------
# 5) Instalar dependencias
# ---------------------------------------------------------
sudo apt install -y build-essential gcc-multilib g++-multilib flex bison || true
sudo apt install -y git wget curl pkg-config gettext || true
sudo apt install -y cups-daemon cups-client printer-driver-all system-config-printer cups-pdf printer-driver-cups-pdf || true
sudo apt install -y msitools || true
sudo apt install -y clang lld || true
sudo apt install -y libc6:i386 libgcc1:i386 libstdc++6:i386 || true
sudo apt install -y libfreetype6:i386 libx11-6:i386 libxext6:i386 libxrender1:i386 libxrandr2:i386 || true
sudo apt install -y winbind samba-common samba-libs gnutls-bin || true
sudo apt install -y ttf-mscorefonts-installer || true
sudo apt install -y wine32:i386 winetricks || true

# ---------------------------------------------------------
# 6) Instalar WineCX
# ---------------------------------------------------------
sudo dpkg -i winecx.deb || true
sudo apt install -f -y || true

# ---------------------------------------------------------
# 7) Copiar prefix Office 365 al HOME
# ---------------------------------------------------------
cp -r .Microsoft_Office_365 "$HOME"

# ---------------------------------------------------------
# 8) Copiar íconos
# ---------------------------------------------------------
sudo mkdir -p /usr/share/icons/hicolor/256x256/apps
sudo cp "$HOME/Descargas/MSO365/Office2016Icons/"*365.svg /usr/share/icons/hicolor/256x256/apps/
sudo gtk-update-icon-cache /usr/share/icons/hicolor/

# ---------------------------------------------------------
# 9) Crear carpeta de lanzadores
# ---------------------------------------------------------
sudo mkdir -p /opt/winecx/launchers
sudo chmod 755 /opt/winecx/launchers

# ---------------------------------------------------------
# 10) Función para crear lanzadores
# ---------------------------------------------------------
create_launcher() {
local name="$1"
local exe="$2"

sudo tee "/opt/winecx/launchers/${name}.sh" > /dev/null <<EOF
#!/bin/bash
set -e
export PATH="/opt/winecx/bin:\$PATH"
export WINEPREFIX="\$HOME/.Microsoft_Office_365"
export LANG=C.UTF-8
export WINEDEBUG=-all

app="C:\\\\Program Files\\\\Microsoft Office\\\\root\\\\Office16\\\\${exe}"
/opt/winecx/bin/wineserver -p >/dev/null 2>&1 || true

if [ \$# -eq 0 ]; then
    exec /opt/winecx/bin/wine "\$app"
else
    for file in "\$@"; do
        fullpath=\$(realpath "\$file")
        winpath="Z:\${fullpath//\//\\\\}"
        /opt/winecx/bin/wine "\$app" "\$winpath"
    done
fi
EOF

sudo chmod +x "/opt/winecx/launchers/${name}.sh"
}

# ---------------------------------------------------------
# 11) Crear lanzadores
# ---------------------------------------------------------

# WORD
create_launcher "word365" "WINWORD.EXE"
sudo tee /usr/share/applications/word365.desktop > /dev/null <<EOF
[Desktop Entry]
Name=Microsoft Word 365
Comment=Procesador de textos de Microsoft Office 365
Exec=/opt/winecx/launchers/word365.sh %F
Type=Application
StartupNotify=true
Terminal=false
Icon=Word365
Categories=Office;WordProcessor;
MimeType=application/msword;application/vnd.openxmlformats-officedocument.wordprocessingml.document;application/vnd.ms-word.document.macroEnabled.12;application/rtf;text/plain;
EOF

# EXCEL
create_launcher "excel365" "EXCEL.EXE"
sudo tee /usr/share/applications/excel365.desktop > /dev/null <<EOF
[Desktop Entry]
Name=Microsoft Excel 365
Comment=Hoja de cálculo de Microsoft Office 365
Exec=/opt/winecx/launchers/excel365.sh %F
Type=Application
StartupNotify=true
Terminal=false
Icon=Excel365
Categories=Office;Spreadsheet;
MimeType=application/vnd.ms-excel;application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;application/vnd.ms-excel.sheet.macroEnabled.12;text/csv;
EOF

# POWERPOINT
create_launcher "powerpoint365" "POWERPNT.EXE"
sudo tee /usr/share/applications/powerpoint365.desktop > /dev/null <<EOF
[Desktop Entry]
Name=Microsoft PowerPoint 365
Comment=Presentaciones de Microsoft Office 365
Exec=/opt/winecx/launchers/powerpoint365.sh %F
Type=Application
StartupNotify=true
Terminal=false
Icon=Powerpoint365
Categories=Office;Presentation;
MimeType=application/vnd.ms-powerpoint;application/vnd.openxmlformats-officedocument.presentationml.presentation;application/vnd.ms-powerpoint.presentation.macroEnabled.12;
EOF

# OUTLOOK
create_launcher "outlook365" "OUTLOOK.EXE"
sudo tee /usr/share/applications/outlook365.desktop > /dev/null <<EOF
[Desktop Entry]
Name=Microsoft Outlook 365
Comment=Cliente de correo de Microsoft Office 365
Exec=/opt/winecx/launchers/outlook365.sh %F
Type=Application
StartupNotify=true
Terminal=false
Icon=Outlook365
Categories=Office;Email;
MimeType=application/vnd.ms-outlook;application/mbox;message/rfc822;
EOF

# ACCESS
create_launcher "access365" "MSACCESS.EXE"
sudo tee /usr/share/applications/access365.desktop > /dev/null <<EOF
[Desktop Entry]
Name=Microsoft Access 365
Comment=Base de datos de Microsoft Office 365
Exec=/opt/winecx/launchers/access365.sh %F
Type=Application
StartupNotify=true
Terminal=false
Icon=Access365
Categories=Office;Database;
MimeType=application/vnd.ms-access;application/x-msaccess;
EOF

# PUBLISHER
create_launcher "publisher365" "MSPUB.EXE"
sudo tee /usr/share/applications/publisher365.desktop > /dev/null <<EOF
[Desktop Entry]
Name=Microsoft Publisher 365
Comment=Publicaciones de Microsoft Office 365
Exec=/opt/winecx/launchers/publisher365.sh %F
Type=Application
StartupNotify=true
Terminal=false
Icon=Publisher365
Categories=Office;Publishing;
MimeType=application/x-mspublisher;
EOF


# ---------------------------------------------------------
# 12) Actualizar base de datos de aplicaciones
# ---------------------------------------------------------
sudo update-desktop-database /usr/share/applications

# ---------------------------------------------------------
# 13) Corregir permisos
# ---------------------------------------------------------
sudo chown -R "$USER:$USER" ~/.Microsoft_Office_365
sudo chmod -R u+rwX ~/.Microsoft_Office_365

# ---------------------------------------------------------
# 14) Reconstruir DOSDEVICES
# ---------------------------------------------------------
rm -rf "$HOME/.Microsoft_Office_365/dosdevices"
mkdir -p "$HOME/.Microsoft_Office_365/dosdevices"

# Unidades principales
ln -s ../drive_c "$HOME/.Microsoft_Office_365/dosdevices/c:"
ln -s / "$HOME/.Microsoft_Office_365/dosdevices/z:"

# Archivos de dispositivo requeridos
ln -s /dev/null "$HOME/.Microsoft_Office_365/dosdevices/c::"
ln -s /dev/null "$HOME/.Microsoft_Office_365/dosdevices/z::"

# (Opcional pero recomendado)
# Lector de CD imaginario
ln -s /media "$HOME/.Microsoft_Office_365/dosdevices/d:"

# Acceso directo al HOME del usuario
ln -s "$HOME" "$HOME/.Microsoft_Office_365/dosdevices/e:"

# ---------------------------------------------------------
# 15) Carpetas de usuario Crossover
# ---------------------------------------------------------
mkdir -p ~/.Microsoft_Office_365/drive_c/users/crossover/AppData/Local
mkdir -p ~/.Microsoft_Office_365/drive_c/users/crossover/AppData/Roaming

# ---------------------------------------------------------
# 16) Actualizar prefix
# ---------------------------------------------------------
WINEPREFIX=$HOME/.Microsoft_Office_365 /opt/winecx/bin/wine wineboot -u

# ---------------------------------------------------------
# 17) Instalar fuentes Winetricks
# ---------------------------------------------------------
#WINEPREFIX=$HOME/.Microsoft_Office_365 winetricks -q corefonts fontsmooth=rgb tahoma fakejapanese

# ---------------------------------------------------------
# 18) Reiniciar servicios wine
# ---------------------------------------------------------
#pkill -9 -f wineserver || true
#pkill -9 -f wine || true
#pkill -9 -f EXCEL.EXE || true
#pkill -9 -f WINWORD.EXE || true
#pkill -9 -f POWERPNT.EXE || true
#pkill -9 -f OUTLOOK.EXE || true
#pkill -9 -f MSACCESS.EXE || true
#pkill -9 -f MSPUB.EXE || true
#pkill -9 -f OFFICEC2RCLIENT.EXE || true
#pkill -9 -f OfficeClickToRun.exe || true

WINEPREFIX=$HOME/.Microsoft_Office_365 /opt/winecx/bin/wineserver -k
WINEPREFIX=$HOME/.Microsoft_Office_365 /opt/winecx/bin/wineserver -w

# ---------------------------------------------------------
# 19) Copiar fuentes de Office 365
# ---------------------------------------------------------
mkdir -p $HOME/.Microsoft_Office_365/drive_c/windows/Fonts
cd "$HOME/Descargas/MSO365/Fuentes Office365"
cp *.ttf *.TTF *.ttc $HOME/.Microsoft_Office_365/drive_c/windows/Fonts

# ---------------------------------------------------------
# 20) Registrar TODAS las fuentes instaladas
# ---------------------------------------------------------
WINEPREFIX=$HOME/.Microsoft_Office_365 bash -c '
FONTDIR="$WINEPREFIX/drive_c/windows/Fonts"
REGFILE="$WINEPREFIX/allfonts.reg"

echo "REGEDIT4" > "$REGFILE"
echo "" >> "$REGFILE"
echo "[HKEY_LOCAL_MACHINE\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Fonts]" >> "$REGFILE"

for f in "$FONTDIR"/*.ttf "$FONTDIR"/*.TTF "$FONTDIR"/*.otf "$FONTDIR"/*.OTF; do
    [ -e "$f" ] || continue
    base=$(basename "$f")
    name="${base%.*}"
    label=$(echo "$name" | sed "s/_/ /g" | sed "s/Regular//g" )

    echo "\"$label (TrueType)\"=\"${base}\"" >> "$REGFILE"
done

/opt/winecx/bin/wine regedit "$REGFILE"
'

echo "==============================================="
echo "        Office 365 Instalado Correctamente"
echo "==============================================="

echo "Recuerda visitarnos en: https://www.youtube.com/@formateando"
sh -c 'xdg-open "https://www.youtube.com/@formateando" >/dev/null 2>&1 &'

