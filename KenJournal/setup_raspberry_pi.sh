#!/bin/bash

# Update and upgrade the system
echo "Updating and upgrading the system..."
sudo apt-get update -y && sudo apt-get upgrade -y

# Install desired packages
echo "Installing necessary packages..."
sudo apt-get install -y arc-theme papirus-icon-theme breeze-cursor-theme syncthing ghostwriter flatpak

# Add Flathub repository and install Apostrophe via Flatpak
echo "Adding Flathub repository and installing Apostrophe..."
sudo flatpak remote-add --if-not-exists flathub https://flathub.org/repo/flathub.flatpakrepo
sudo flatpak install -y flathub org.gnome.Apostrophe

# Remove unwanted packages
echo "Uninstalling unnecessary packages..."
sudo apt-get remove -y vlc

# Install Python packages (if needed)
echo "Installing Python packages..."
sudo apt-get install -y python3-pip
# Uncomment the next line to install jrnl
# pip3 install jrnl

# Set GTK theme
echo "Setting GTK theme..."
mkdir -p ~/.config/gtk-3.0/
cat <<EOF > ~/.config/gtk-3.0/settings.ini
[Settings]
gtk-theme-name=Arc-Dark
gtk-icon-theme-name=Papirus
gtk-cursor-theme-name=Breeze
EOF

# Set icon theme
echo "Setting icon theme..."
mkdir -p ~/.icons/default/
cat <<EOF > ~/.icons/default/index.theme
[Icon Theme]
Inherits=Papirus
EOF

# Set cursor theme
echo "Setting cursor theme..."
cat <<EOF > ~/.icons/default/index.theme
[Icon Theme]
Inherits=Breeze
EOF

# Notify user of success
echo "Themes installed and set successfully!"

# Clean up
echo "Cleaning up..."
sudo apt-get autoremove -y
sudo apt-get clean

# Reboot the system
echo "Rebooting the system..."
sudo reboot
