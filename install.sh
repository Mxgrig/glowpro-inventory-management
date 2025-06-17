#!/bin/bash

# GlowPro Inventory Management System - Automated Installer
# Copyright (c) 2024 GlowPro Solutions
# 
# One-click installation script for the Beauty Pro Inventory System
# Usage: curl -sSL https://raw.githubusercontent.com/Mxgrig/glowpro-inventory-management/main/install.sh | bash

set -e

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
PURPLE='\033[0;35m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

# Configuration
REPO_URL="https://github.com/Mxgrig/glowpro-inventory-management"
RAW_URL="https://raw.githubusercontent.com/Mxgrig/glowpro-inventory-management/main"
INSTALL_DIR="$HOME/GlowPro-Inventory"
TEMP_DIR="/tmp/glowpro-install-$$"

# Helper functions
print_header() {
    echo -e "${PURPLE}"
    echo "╔═══════════════════════════════════════════════════════════════╗"
    echo "║                 💄 GlowPro Inventory System                   ║"
    echo "║              Professional Beauty Business Solution             ║"
    echo "║                        Automated Installer                    ║"
    echo "╚═══════════════════════════════════════════════════════════════╝"
    echo -e "${NC}"
}

print_step() {
    echo -e "${CYAN}[STEP]${NC} $1"
}

print_success() {
    echo -e "${GREEN}[SUCCESS]${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

print_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

print_info() {
    echo -e "${BLUE}[INFO]${NC} $1"
}

# Check system requirements
check_requirements() {
    print_step "Checking system requirements..."
    
    # Check for curl
    if ! command -v curl &> /dev/null; then
        print_error "curl is required but not installed. Please install curl and try again."
        exit 1
    fi
    
    # Check for unzip (for potential future use)
    if ! command -v unzip &> /dev/null; then
        print_warning "unzip not found. Some features may not work properly."
    fi
    
    # Check internet connection
    if ! curl -s --head https://google.com &> /dev/null; then
        print_error "Internet connection required. Please check your connection and try again."
        exit 1
    fi
    
    print_success "System requirements check passed"
}

# Create installation directory
create_install_dir() {
    print_step "Creating installation directory..."
    
    if [ -d "$INSTALL_DIR" ]; then
        print_warning "Installation directory already exists: $INSTALL_DIR"
        echo -n "Remove existing installation? (y/N): "
        read -r response
        if [[ "$response" =~ ^[Yy]$ ]]; then
            rm -rf "$INSTALL_DIR"
            print_info "Removed existing installation"
        else
            print_error "Installation cancelled"
            exit 1
        fi
    fi
    
    mkdir -p "$INSTALL_DIR"
    mkdir -p "$TEMP_DIR"
    
    print_success "Created installation directory: $INSTALL_DIR"
}

# Download files
download_files() {
    print_step "Downloading GlowPro Inventory System files..."
    
    cd "$TEMP_DIR"
    
    # List of files to download
    files=(
        "README.md"
        "SETUP_GUIDE.md"
        "AUTO_INSTALLER.md"
        "EXCEL_WORKBOOK_README.md"
        "CHANGELOG.md"
        "Beauty_Pro_Inventory_System_FINAL.xlsx"
        "sheets/Dashboard.csv"
        "sheets/Categories.csv"
        "sheets/Suppliers.csv"
        "sheets/Products.csv"
        "sheets/Inventory.csv"
        "sheets/QuickAdd.csv"
        "sheets/Reorder.csv"
        "sheets/Analytics.csv"
        "sheets/Instructions.csv"
        "scripts/Code.gs"
        "scripts/Setup.gs"
        "scripts/Validation.gs"
        "assets/colors.css"
    )
    
    # Create subdirectories
    mkdir -p sheets scripts assets
    
    # Download each file
    for file in "${files[@]}"; do
        print_info "Downloading $file..."
        if curl -sSL "$RAW_URL/$file" -o "$file"; then
            echo -e "  ${GREEN}✓${NC} $file"
        else
            print_warning "Failed to download $file (continuing...)"
        fi
    done
    
    print_success "File download completed"
}

# Install files
install_files() {
    print_step "Installing GlowPro Inventory System..."
    
    # Copy files to installation directory
    cp -r "$TEMP_DIR"/* "$INSTALL_DIR/"
    
    # Make sure permissions are correct
    chmod -R 755 "$INSTALL_DIR"
    
    print_success "Installation completed"
}

# Create desktop shortcut (Linux)
create_shortcuts() {
    print_step "Creating shortcuts..."
    
    # Create a simple launcher script
    cat > "$INSTALL_DIR/launch.sh" << 'EOF'
#!/bin/bash
# GlowPro Inventory System Launcher

echo "🚀 GlowPro Inventory Management System"
echo "======================================"
echo ""
echo "Choose your installation method:"
echo ""
echo "1. Open Google Sheets Template (Recommended)"
echo "2. Open Excel File Locally"
echo "3. View Setup Guide"
echo "4. View Documentation"
echo ""
echo -n "Enter your choice (1-4): "
read choice

case $choice in
    1)
        echo ""
        echo "📋 Google Sheets Installation:"
        echo "1. Go to: https://sheets.google.com"
        echo "2. Create a new spreadsheet"
        echo "3. Follow the SETUP_GUIDE.md instructions"
        echo ""
        if command -v xdg-open &> /dev/null; then
            xdg-open "https://sheets.google.com"
        fi
        ;;
    2)
        echo ""
        echo "📊 Opening Excel file..."
        if command -v xdg-open &> /dev/null; then
            xdg-open "Beauty_Pro_Inventory_System_FINAL.xlsx"
        else
            echo "Please open: Beauty_Pro_Inventory_System_FINAL.xlsx"
        fi
        ;;
    3)
        if command -v xdg-open &> /dev/null; then
            xdg-open "SETUP_GUIDE.md"
        else
            echo "Please open: SETUP_GUIDE.md"
        fi
        ;;
    4)
        if command -v xdg-open &> /dev/null; then
            xdg-open "README.md"
        else
            echo "Please open: README.md"
        fi
        ;;
    *)
        echo "Invalid choice. Please run the script again."
        ;;
esac
EOF
    
    chmod +x "$INSTALL_DIR/launch.sh"
    
    # Try to create desktop shortcut (Linux with desktop environment)
    if [ -d "$HOME/Desktop" ] && command -v xdg-user-dir &> /dev/null; then
        cat > "$HOME/Desktop/GlowPro-Inventory.desktop" << EOF
[Desktop Entry]
Version=1.0
Type=Application
Name=GlowPro Inventory System
Comment=Professional Beauty Business Inventory Management
Exec=$INSTALL_DIR/launch.sh
Icon=$INSTALL_DIR/assets/icon.png
Terminal=true
Categories=Office;Finance;
EOF
        chmod +x "$HOME/Desktop/GlowPro-Inventory.desktop"
        print_success "Created desktop shortcut"
    fi
    
    print_success "Shortcuts created"
}

# Display installation summary
show_summary() {
    echo ""
    echo -e "${PURPLE}╔═══════════════════════════════════════════════════════════════╗${NC}"
    echo -e "${PURPLE}║                    🎉 Installation Complete!                  ║${NC}"
    echo -e "${PURPLE}╚═══════════════════════════════════════════════════════════════╝${NC}"
    echo ""
    echo -e "${GREEN}✅ GlowPro Inventory System installed successfully!${NC}"
    echo ""
    echo -e "${CYAN}📁 Installation Location:${NC} $INSTALL_DIR"
    echo ""
    echo -e "${YELLOW}🚀 Next Steps:${NC}"
    echo "1. Run the launcher: $INSTALL_DIR/launch.sh"
    echo "2. Choose Google Sheets setup (recommended)"
    echo "3. Follow the guided setup process"
    echo "4. Start managing your beauty business inventory!"
    echo ""
    echo -e "${BLUE}📚 Documentation:${NC}"
    echo "• Setup Guide: $INSTALL_DIR/SETUP_GUIDE.md"
    echo "• User Manual: $INSTALL_DIR/README.md"
    echo "• Auto Installer Guide: $INSTALL_DIR/AUTO_INSTALLER.md"
    echo ""
    echo -e "${PURPLE}💄 Transform your beauty business with professional inventory management!${NC}"
    echo ""
    
    # Offer to launch immediately
    echo -n "Would you like to launch GlowPro now? (Y/n): "
    read -r response
    if [[ ! "$response" =~ ^[Nn]$ ]]; then
        cd "$INSTALL_DIR"
        ./launch.sh
    fi
}

# Cleanup function
cleanup() {
    if [ -d "$TEMP_DIR" ]; then
        rm -rf "$TEMP_DIR"
    fi
}

# Error handling
error_handler() {
    print_error "Installation failed on line $1"
    cleanup
    exit 1
}

# Set error trap
trap 'error_handler $LINENO' ERR

# Main installation process
main() {
    print_header
    
    print_info "Starting GlowPro Inventory System installation..."
    echo ""
    
    check_requirements
    create_install_dir
    download_files
    install_files
    create_shortcuts
    
    cleanup
    
    show_summary
}

# Run main function
main "$@"