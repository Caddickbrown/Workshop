#!/bin/bash

# Theming

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

# Final Clean up and Reboot

# Clean up
echo "Cleaning up..."
sudo apt-get autoremove -y
sudo apt-get clean

# Reboot the system
echo "Rebooting the system..."
sudo reboot
