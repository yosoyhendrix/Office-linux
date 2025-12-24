#!/bin/bash

# ==============================================================================
# SCRIPT DE INSTALACIÓN AUTOMATIZADA DE OFFICE 365 PARA LINUX (Ubuntu/Debian)
# Basado en las instrucciones proporcionadas en el https://youtu.be/jP0y5Pdq2HU?si=d_5_3EojYOXrMsz7
# ==============================================================================

set -e # Detener el script si ocurre un error

# Definir colores para mensajes
GREEN='\033[0;32m'
BLUE='\033[0;34m'
RED='\033[0;31m'
NC='\033[0m' # No Color

echo -e "${BLUE}>>> Iniciando script de instalación de Office 365...${NC}"

# 1. Verificar directorios y archivo ZIP
TARGET_DIR="$HOME/Descargas"
if [ ! -d "$TARGET_DIR" ]; then
    TARGET_DIR="$HOME/Downloads"
fi

ZIP_FILE="$TARGET_DIR/MSO365.zip"

if [ ! -f "$ZIP_FILE" ]; then
    echo -e "${RED}Error: No se encontró el archivo MSO365.zip en $TARGET_DIR${NC}"
    exit 1
fi

EXTRACT_DIR="$TARGET_DIR/MSO365"

echo -e "${GREEN}>>> 1. Descomprimiendo archivos...${NC}"
cd "$TARGET_DIR"
unzip -o "$ZIP_FILE"

# Entrar a la carpeta descomprimida
if [ -d "$EXTRACT_DIR" ]; then
    cd "$EXTRACT_DIR"
else
    echo -e "${RED}Error: No se encontró la carpeta descomprimida MSO365.${NC}"
    exit 1
fi

echo -e "${GREEN}>>> 2. Habilitando arquitectura de 32 bits y actualizando sistema...${NC}"
sudo dpkg --add-architecture i386
sudo apt update

echo -e "${GREEN}>>> 3. Instalando dependencias necesarias...${NC}"
# Se agrupan las instalaciones para eficiencia, usando || true como en el manual para paquetes opcionales
sudo apt install -y build-essential gcc-multilib g++-multilib flex bison || true
sudo apt install -y git wget curl pkg-config gettext || true
sudo apt install -y cups-daemon cups-client printer-driver-all system-config-printer cups-pdf printer-driver-cups-pdf || true
sudo apt install -y msitools || true
sudo apt install -y clang lld || true
sudo apt install -y libc6:i386 libgcc1:i386 libstdc++6:i386 || true
sudo apt install -y libfreetype6:i386 libx11-6:i386 libxext6:i386 libxrender1:i386 libxrandr2:i386 || true
sudo apt install -y libfreetype6:i386 || true
sudo apt install -y winbind samba-common samba-libs gnutls-bin || true
sudo apt install -y ttf-mscorefonts-installer || true
sudo apt install -y wine32:i386 || true
sudo apt install -y winetricks || true

echo -e "${GREEN}>>> 4. Instalando WineCX...${NC}"
# Buscamos winecx.deb dentro de la carpeta actual (MSO365)
if [ -f "winecx.deb" ]; then
    sudo dpkg -i winecx.deb || true
    sudo apt install -f -y || true # Corregir dependencias rotas si las hay
else
    echo -e "${RED}Alerta: No se encontró winecx.deb en la carpeta MSO365. Intentando continuar...${NC}"
fi

echo -e "${GREEN}>>> 5. Configurando el Prefix de Wine...${NC}"
# Copiar Prefix .Microsoft_Office_365 al Home
if [ -d ".Microsoft_Office_365" ]; then
    cp -r ".Microsoft_Office_365" "$HOME/"
else
    echo -e "${RED}Error: No se encontró la carpeta .Microsoft_Office_365 dentro del zip.${NC}"
    exit 1
fi

echo -e "${GREEN}>>> 6. Instalando Iconos...${NC}"
sudo mkdir -p /usr/share/icons/hicolor/256x256/apps
# Ajuste de ruta basado en PDF
sudo cp "$EXTRACT_DIR/Office2016Icons/"*365.svg /usr/share/icons/hicolor/256x256/apps/ || true
sudo gtk-update-icon-cache /usr/share/icons/hicolor/ || true

echo -e "${GREEN}>>> 7. Creando lanzadores y scripts...${NC}"
sudo mkdir -p /opt/winecx/launchers
sudo chmod 755 /opt/winecx/launchers

# Función para crear el script lanzador en /opt/winecx
create_launcher_script() {
    local name="$1"
    local exe="$2"
    
    # Creamos el archivo temporalmente y luego usamos sudo move/tee
    cat <<EOF > "/tmp/${name}.sh"
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
    sudo mv "/tmp/${name}.sh" "/opt/winecx/launchers/${name}.sh"
    sudo chmod +x "/opt/winecx/launchers/${name}.sh"
}

# Crear scripts internos
create_launcher_script "word365" "WINWORD.EXE"
create_launcher_script "excel365" "EXCEL.EXE"
create_launcher_script "powerpoint365" "POWERPNT.EXE"
create_launcher_script "outlook365" "OUTLOOK.EXE"
create_launcher_script "access365" "MSACCESS.EXE"
create_launcher_script "publisher365" "MSPUB.EXE"

echo -e "${GREEN}>>> 8. Creando accesos directos (.desktop)...${NC}"

# Word
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

# Excel
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

# PowerPoint
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

# Outlook
sudo tee /usr/share/applications/outlook365.desktop > /dev/null <<EOF
[Desktop Entry]
Name=Microsoft Outlook 365
Comment=Cliente de correo electrónico de Microsoft Office 365
Exec=/opt/winecx/launchers/outlook365.sh %F
Type=Application
StartupNotify=true
Terminal=false
Icon=Outlook365
Categories=Office;Email;
MimeType=application/vnd.ms-outlook;application/mbox;message/rfc822;
EOF

# Access
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

# Publisher
sudo tee /usr/share/applications/publisher365.desktop > /dev/null <<EOF
[Desktop Entry]
Name=Microsoft Publisher 365
Comment=Editor de publicaciones de Microsoft Office 365
Exec=/opt/winecx/launchers/publisher365.sh %F
Type=Application
StartupNotify=true
Terminal=false
Icon=Publisher365
Categories=Office;Publishing;
MimeType=application/x-mspublisher;
EOF

sudo update-desktop-database /usr/share/applications

echo -e "${GREEN}>>> 9. Configurando permisos y directorios de usuario...${NC}"
sudo chown -R "$USER:$USER" "$HOME/.Microsoft_Office_365"
sudo chmod -R u+rwX "$HOME/.Microsoft_Office_365"

# Reconstruir enlaces simbólicos
rm -rf "$HOME/.Microsoft_Office_365/dosdevices"
mkdir -p "$HOME/.Microsoft_Office_365/dosdevices"
ln -s "../drive_c" "$HOME/.Microsoft_Office_365/dosdevices/c:"
ln -s "/" "$HOME/.Microsoft_Office_365/dosdevices/z:"

# Crear carpetas de usuario dentro de Wine
mkdir -p "$HOME/.Microsoft_Office_365/drive_c/users/crossover/AppData/Local"
mkdir -p "$HOME/.Microsoft_Office_365/drive_c/users/crossover/AppData/Roaming"

echo -e "${GREEN}>>> 10. Actualizando Prefix e instalando fuentes base...${NC}"
export WINEPREFIX="$HOME/.Microsoft_Office_365"

# Ejecutar wineboot
/opt/winecx/bin/wine wineboot -u || true

# Winetricks (Silencioso)
winetricks -q corefonts || true
winetricks -q fontsmooth=rgb || true
winetricks -q fontsmooth-gray || true
winetricks -q tahoma || true
winetricks -q fakejapanese || true

# Reiniciar servicios de Wine
echo "Reiniciando Wine..."
pkill -9 wineserver || true
pkill -9 wine || true
pkill -9 EXCEL.EXE || true
pkill -9 WINWORD.EXE || true
pkill -9 OFFICEC2RCLIENT.EXE || true
pkill -9 OfficeClickToRun.exe || true

/opt/winecx/bin/wineserver -k || true
/opt/winecx/bin/wineserver -w || true

echo -e "${GREEN}>>> 11. Copiando y registrando fuentes de Office 365...${NC}"
mkdir -p "$HOME/.Microsoft_Office_365/drive_c/windows/Fonts"

# Copiar fuentes desde carpeta descomprimida (ajustar nombre carpeta si difiere)
if [ -d "$EXTRACT_DIR/Fuentes Office365" ]; then
    cd "$EXTRACT_DIR/Fuentes Office365"
    cp *.ttf *.TTF *.ttc "$HOME/.Microsoft_Office_365/drive_c/windows/Fonts" || true
else
    echo -e "${RED}Advertencia: No se encontró carpeta 'Fuentes Office365'. Saltando copia de fuentes.${NC}"
fi

# Script para registrar fuentes en el registro de Wine
# (Lógica extraída directamente de la página 4 del PDF)
echo "Generando archivo de registro de fuentes..."

FONTDIR="$WINEPREFIX/drive_c/windows/Fonts"
REGFILE="$WINEPREFIX/allfonts.reg"

echo "REGEDIT4" > "$REGFILE"
echo "" >> "$REGFILE"
echo "[HKEY_LOCAL_MACHINE\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Fonts]" >> "$REGFILE"

# Iterar sobre las fuentes y añadirlas al registro
for f in "$FONTDIR"/*.ttf "$FONTDIR"/*.TTF "$FONTDIR"/*.otf "$FONTDIR"/*.OTF; do
    [ -e "$f" ] || continue
    base=$(basename "$f")
    name="${base%.*}"
    # Convertir nombre a formato amigable
    label=$(echo "$name" | sed "s/ //g" | sed "s/Regular//g")
    echo "\"$label (TrueType)\"=\"$base\"" >> "$REGFILE"
done

# Importar el registro
/opt/winecx/bin/wine regedit "$REGFILE"

echo -e "${GREEN}======================================================${NC}"
echo -e "${GREEN}   INSTALACIÓN COMPLETADA EXITOSAMENTE    ${NC}"
echo -e "${GREEN}   Busca 'Word' o 'Excel' en tu menú de aplicaciones.${NC}"
echo -e "${GREEN}======================================================${NC}"
