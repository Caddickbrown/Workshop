#!/bin/bash

# Download the JetBrains Mono font
cd /tmp
wget https://github.com/JetBrains/JetBrainsMono/releases/download/v2.242/JetBrainsMono-2.242.zip

# Extract the font files
unzip JetBrainsMono-2.242.zip

# Install the font
sudo mv JetBrainsMono-2.242/ttf/*.ttf /usr/share/fonts/
sudo fc-cache -f -v

# Set the font as default for GNOME Terminal
gsettings set org.gnome.desktop.interface monospace-font-name 'JetBrains Mono 12'

# Set the font as default for Raspberry Pi OS Terminal
if grep -q "Raspberry Pi" /etc/os-release; then
    echo "Setting JetBrains Mono as default for Raspberry Pi OS Terminal"
    sudo sh -c 'echo "TERM_FONT=JetBrainsMono" >> /etc/environment'
fi

echo "JetBrains Mono font installed and set as default."