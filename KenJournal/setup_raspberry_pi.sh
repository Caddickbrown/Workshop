#!/bin/bash

# Package Setup

# Update and upgrade the system
echo "Updating and upgrading the system..."
sudo apt-get update -y && sudo apt-get upgrade -y

# Install desired packages
echo "Installing necessary packages..."
sudo apt-get install -y arc-theme papirus-icon-theme breeze-cursor-theme syncthing ghostwriter flatpak chromium-browser python3-pip

# Install Homebrew Packages
#/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
#brew install jrnl

# Add Flathub repository and install Apostrophe via Flatpak
echo "Adding Flathub repository and installing Apostrophe..."
sudo flatpak remote-add --if-not-exists flathub https://flathub.org/repo/flathub.flatpakrepo
sudo flatpak install -y flathub org.gnome.gitlab.somas.Apostrophe

# Remove unwanted packages
echo "Uninstalling unnecessary packages..."
sudo apt-get remove -y vlc

# Install Python packages (if needed)
#echo "Installing Python packages..."

# Display Settings

# Set the default resolution to 720p and refresh rate to 30Hz
#echo "Setting default resolution to 720p and refresh rate to 30Hz..."
#sudo sed -i '/^hdmi_group=/d' /boot/config.txt
#sudo sed -i '/^hdmi_mode=/d' /boot/config.txt
#sudo sed -i '/^hdmi_force_hotplug=/d' /boot/config.txt

#echo "hdmi_group=1" | sudo tee -a /boot/config.txt
#echo "hdmi_mode=4" | sudo tee -a /boot/config.txt
#echo "hdmi_force_hotplug=1" | sudo tee -a /boot/config.txt

# Explanation:
# - hdmi_group=1: CEA (Consumer Electronics Association) group, used for TVs.
# - hdmi_mode=4: 720p resolution at 60Hz (by default). We need to override this to 30Hz.
# - hdmi_force_hotplug=1: Forces the HDMI mode even if no HDMI monitor is detected.

# Set the refresh rate to 30Hz
#echo "Setting refresh rate to 30Hz..."
#sudo sed -i '/^hdmi_mode/!b;n;c\hdmi_mode=2' /boot/config.txt
#hdmi_mode=2 corresponds to 720p at 30Hz.

# Theming

# Set GTK theme
#echo "Setting GTK theme..."
#mkdir -p ~/.config/gtk-3.0/
#cat <<EOF > ~/.config/gtk-3.0/settings.ini
#[Settings]
#gtk-theme-name=Arc-Dark
#gtk-icon-theme-name=Papirus
#gtk-cursor-theme-name=Breeze
#EOF

# Set icon theme
#echo "Setting icon theme..."
#mkdir -p ~/.icons/default/
#cat <<EOF > ~/.icons/default/index.theme
#[Icon Theme]
#Inherits=Papirus
#EOF

# Set cursor theme
#echo "Setting cursor theme..."
#cat <<EOF > ~/.icons/default/index.theme
#[Icon Theme]
#Inherits=Breeze
#EOF

# Notify user of success
#echo "Themes installed and set successfully!"

# Define the configuration file location
#CONFIG_DIR="$HOME/.config/lxpanel/LXDE-pi/panels"
#CONFIG_FILE="$CONFIG_DIR/panel"

# Ensure the LXPanel configuration directory exists
#if [ ! -d "$CONFIG_DIR" ]; then
#    echo "Creating LXPanel config directory..."
#    mkdir -p "$CONFIG_DIR"
#fi

# Create a basic configuration file if it doesn't exist
#if [ ! -f "$CONFIG_FILE" ]; then
#    echo "Creating basic panel configuration file..."
#    cat <<EOF > "$CONFIG_FILE"
#Global {
#    edge=bottom
#    autohide=true
#}
#EOF
#fi

# Set the panel to the left side of the screen and enable auto-hide
#echo "Setting panel to the left side and enabling auto-hide..."
#sed -i 's/^edge=.*/edge=left/' "$CONFIG_FILE"
#sed -i 's/^autohide=.*/autohide=true/' "$CONFIG_FILE"

# Restart the LXPanel to apply changes
#echo "Restarting LXPanel..."
#lxpanelctl restart

# Final Clean up and Reboot

# Clean up
echo "Cleaning up..."
sudo apt-get autoremove -y
sudo apt-get clean

# Reboot the system
echo "Rebooting the system..."
sudo reboot
