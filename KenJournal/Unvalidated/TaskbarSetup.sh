#!/bin/bash

# Taskbar Setup

# Define the configuration file location
CONFIG_DIR="$HOME/.config/lxpanel/LXDE-pi/panels"
CONFIG_FILE="$CONFIG_DIR/panel"

# Ensure the LXPanel configuration directory exists
if [ ! -d "$CONFIG_DIR" ]; then
    echo "Creating LXPanel config directory..."
    mkdir -p "$CONFIG_DIR"
fi

# Create a basic configuration file if it doesn't exist
if [ ! -f "$CONFIG_FILE" ]; then
    echo "Creating basic panel configuration file..."
    cat <<EOF > "$CONFIG_FILE"
Global {
    edge=bottom
    autohide=true
}
EOF
fi

# Set the panel to the left side of the screen and enable auto-hide
echo "Setting panel to the left side and enabling auto-hide..."
sed -i 's/^edge=.*/edge=left/' "$CONFIG_FILE"
sed -i 's/^autohide=.*/autohide=true/' "$CONFIG_FILE"

# Restart the LXPanel to apply changes
echo "Restarting LXPanel..."
lxpanelctl restart

# Final Clean up and Reboot

# Clean up
echo "Cleaning up..."
sudo apt-get autoremove -y
sudo apt-get clean

# Reboot the system
echo "Rebooting the system..."
sudo reboot
