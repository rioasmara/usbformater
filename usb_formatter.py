import sys
import os
import subprocess
import string
import time
from queue import Queue
from threading import Thread, Lock
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QTextEdit, QLabel,
                             QComboBox, QMessageBox, QGroupBox, QTableWidget,
                             QTableWidgetItem, QHeaderView, QCheckBox, QProgressBar)
from PyQt6.QtCore import QTimer, Qt, pyqtSignal, QObject
from PyQt6.QtGui import QFont, QColor
import win32api
import win32file
import win32con
import ctypes
import wmi
from win32api import GetVolumeInformation
from win32file import GetDriveType
import win32com.client


class USBMonitor(QObject):
    """Monitor for USB device changes including USB and portable SCSI drives"""
    usb_connected = pyqtSignal(str)  # Signal emitted when USB is connected
    usb_disconnected = pyqtSignal(str)  # Signal emitted when USB is disconnected

    def __init__(self):
        super().__init__()
        self.wmi_client = wmi.WMI()
        self.previous_drives = set(self.get_removable_drives())

    def is_external_drive(self, drive_letter, debug=False):
        """Check if drive is external (USB or portable SCSI with surprise removal policy)"""
        try:
            for physical_disk in self.wmi_client.Win32_DiskDrive():
                for partition in physical_disk.associators("Win32_DiskDriveToDiskPartition"):
                    for logical_disk in partition.associators("Win32_LogicalDiskToPartition"):
                        if logical_disk.DeviceID == f"{drive_letter}:":
                            interface_type = physical_disk.InterfaceType or ""
                            pnp_id = physical_disk.PNPDeviceID or ""
                            media_type = physical_disk.MediaType or ""

                            if debug:
                                print(f"[DEBUG] Drive {drive_letter}: InterfaceType='{interface_type}'")
                                print(f"[DEBUG] Drive {drive_letter}: PNP_ID='{pnp_id}'")
                                print(f"[DEBUG] Drive {drive_letter}: MediaType='{media_type}'")

                            # Check if USB is anywhere in the device tree/parent devices
                            is_usb_device = self.check_if_usb_in_device_tree(pnp_id, debug)
                            if is_usb_device:
                                if debug:
                                    print(f"[DEBUG] Drive {drive_letter}: Found USB in device tree - ACCEPTED")
                                return True

                            # Direct USB interface check
                            if "USB" in interface_type.upper() or "USB" in pnp_id.upper() or "USBSTOR" in pnp_id.upper():
                                if debug:
                                    print(f"[DEBUG] Drive {drive_letter}: Detected as USB device - ACCEPTED")
                                return True

                            # Check for removable or external media type
                            if "REMOVABLE" in media_type.upper():
                                if debug:
                                    print(f"[DEBUG] Drive {drive_letter}: Has REMOVABLE media type - ACCEPTED")
                                return True

                            # Check for external media type (USB drives often report as "External hard disk media")
                            if "EXTERNAL" in media_type.upper():
                                if debug:
                                    print(f"[DEBUG] Drive {drive_letter}: Has EXTERNAL media type - ACCEPTED")
                                return True

                            # For SCSI, check if it has surprise removal policy (portable)
                            if "SCSI" in interface_type.upper():
                                if debug:
                                    print(f"[DEBUG] Drive {drive_letter}: Detected as SCSI, checking removal policy...")
                                if self.is_portable_scsi(pnp_id, debug):
                                    if debug:
                                        print(f"[DEBUG] Drive {drive_letter}: SCSI device has surprise removal policy - ACCEPTED")
                                    return True
                                else:
                                    if debug:
                                        print(f"[DEBUG] Drive {drive_letter}: SCSI device does NOT have surprise removal policy - REJECTED")
                                    return False

                            if debug:
                                print(f"[DEBUG] Drive {drive_letter}: Not recognized as external drive - REJECTED")
                            return False

            if debug:
                print(f"[DEBUG] Drive {drive_letter}: No WMI disk information found")
            return False
        except Exception as e:
            if debug:
                print(f"[DEBUG] Error checking drive {drive_letter}: {e}")
            return False

    def check_if_usb_in_device_tree(self, pnp_device_id, debug=False):
        """Check if USB appears anywhere in the device's parent tree"""
        try:
            import win32com.client
            wmi_obj = win32com.client.GetObject("winmgmts:")

            # Query for the device
            query = f"SELECT * FROM Win32_PnPEntity WHERE DeviceID = '{pnp_device_id.replace(chr(92), chr(92)+chr(92))}'"
            devices = wmi_obj.ExecQuery(query)

            for device in devices:
                # Check parent devices
                try:
                    parent_id = device.Properties_("PNPDeviceID").Value
                    if parent_id and ("USB" in parent_id.upper() or "USBSTOR" in parent_id.upper()):
                        if debug:
                            print(f"[DEBUG] Found USB in device PNP ID: {parent_id}")
                        return True
                except:
                    pass

            return False
        except Exception as e:
            if debug:
                print(f"[DEBUG] Error checking device tree: {e}")
            return False

    def is_portable_scsi(self, pnp_device_id, debug=False):
        """Check if SCSI device has CM_REMOVAL_POLICY_EXPECT_SURPRISE_REMOVAL policy"""
        try:
            if debug:
                print(f"[DEBUG] Checking removal policy for: {pnp_device_id}")
            # CM_REMOVAL_POLICY_EXPECT_SURPRISE_REMOVAL = 3
            CM_REMOVAL_POLICY_EXPECT_SURPRISE_REMOVAL = 3
            SPDRP_REMOVAL_POLICY = 0x0000001F  # Device removal policy property

            # Use SetupAPI to query device properties
            setupapi = ctypes.windll.setupapi

            # Get device info set
            DIGCF_PRESENT = 0x00000002
            DIGCF_ALLCLASSES = 0x00000004

            class SP_DEVINFO_DATA(ctypes.Structure):
                _fields_ = [
                    ("cbSize", ctypes.c_ulong),
                    ("ClassGuid", ctypes.c_byte * 16),
                    ("DevInst", ctypes.c_ulong),
                    ("Reserved", ctypes.c_void_p)
                ]

            h_dev_info = setupapi.SetupDiGetClassDevsW(
                None,
                None,
                None,
                DIGCF_PRESENT | DIGCF_ALLCLASSES
            )

            if h_dev_info == -1:
                return False

            try:
                dev_info_data = SP_DEVINFO_DATA()
                dev_info_data.cbSize = ctypes.sizeof(SP_DEVINFO_DATA)

                index = 0
                while setupapi.SetupDiEnumDeviceInfo(h_dev_info, index, ctypes.byref(dev_info_data)):
                    index += 1

                    # Get device instance ID
                    buffer_size = ctypes.c_ulong(0)
                    setupapi.SetupDiGetDeviceInstanceIdW(
                        h_dev_info,
                        ctypes.byref(dev_info_data),
                        None,
                        0,
                        ctypes.byref(buffer_size)
                    )

                    if buffer_size.value == 0:
                        continue

                    instance_id = ctypes.create_unicode_buffer(buffer_size.value)
                    if not setupapi.SetupDiGetDeviceInstanceIdW(
                        h_dev_info,
                        ctypes.byref(dev_info_data),
                        instance_id,
                        buffer_size,
                        None
                    ):
                        continue

                    # Check if this is our device
                    if pnp_device_id.upper() in instance_id.value.upper():
                        if debug:
                            print(f"[DEBUG] Found matching device: {instance_id.value}")

                        # Get removal policy
                        removal_policy = ctypes.c_ulong(0)
                        property_type = ctypes.c_ulong(0)

                        result = setupapi.SetupDiGetDeviceRegistryPropertyW(
                            h_dev_info,
                            ctypes.byref(dev_info_data),
                            SPDRP_REMOVAL_POLICY,
                            ctypes.byref(property_type),
                            ctypes.byref(removal_policy),
                            ctypes.sizeof(removal_policy),
                            None
                        )

                        if result:
                            if debug:
                                print(f"[DEBUG] Removal policy value: {removal_policy.value} (3=surprise removal)")
                            # Check if removal policy is CM_REMOVAL_POLICY_EXPECT_SURPRISE_REMOVAL
                            is_surprise_removal = removal_policy.value == CM_REMOVAL_POLICY_EXPECT_SURPRISE_REMOVAL
                            if debug:
                                print(f"[DEBUG] Is surprise removal policy: {is_surprise_removal}")
                            return is_surprise_removal
                        else:
                            if debug:
                                print(f"[DEBUG] Could not get removal policy property")

                if debug:
                    print(f"[DEBUG] Device not found in enumeration")
                return False

            finally:
                setupapi.SetupDiDestroyDeviceInfoList(h_dev_info)

        except Exception as e:
            # If we can't determine, exclude it for safety
            return False

    def get_removable_drives(self, debug=False):
        """Get list of removable drives (USB and portable SCSI drives with surprise removal policy)"""
        drives = []
        bitmask = win32api.GetLogicalDrives()
        for letter in string.ascii_uppercase:
            if bitmask & 1:
                drive_path = f"{letter}:\\"
                try:
                    drive_type = win32file.GetDriveType(drive_path)
                    if debug:
                        print(f"[DEBUG] Drive {letter}: type={drive_type} (2=REMOVABLE, 3=FIXED, 5=CDROM)")

                    # DRIVE_REMOVABLE = 2, DRIVE_FIXED = 3, DRIVE_CDROM = 5
                    # Always include DRIVE_REMOVABLE (traditional USB flash drives)
                    if drive_type == win32con.DRIVE_REMOVABLE:
                        if debug:
                            print(f"[DEBUG] Drive {letter}: Is DRIVE_REMOVABLE, checking details...")
                        # Check if it's actually USB or external
                        result = self.is_external_drive(letter, debug)
                        if result:
                            if debug:
                                print(f"[DEBUG] Drive {letter}: Confirmed as external, adding to list")
                            drives.append(letter)
                        else:
                            # Even if not confirmed, add DRIVE_REMOVABLE anyway (it's usually USB)
                            if debug:
                                print(f"[DEBUG] Drive {letter}: Could not confirm interface but is REMOVABLE, adding anyway")
                            drives.append(letter)
                    # Also check fixed drives that might be external (SCSI, USB HDD)
                    elif drive_type == win32con.DRIVE_FIXED:
                        if debug:
                            print(f"[DEBUG] Drive {letter}: Is DRIVE_FIXED, checking if external...")
                        # Use WMI to verify if it's external
                        if self.is_external_drive(letter, debug):
                            if debug:
                                print(f"[DEBUG] Drive {letter}: Verified as external, adding to list")
                            drives.append(letter)
                        else:
                            if debug:
                                print(f"[DEBUG] Drive {letter}: Not external, skipping")
                except Exception as e:
                    if debug:
                        print(f"[DEBUG] Drive {letter}: Exception - {e}")
                    pass
            bitmask >>= 1

        if debug:
            print(f"[DEBUG] Final drives list: {drives}")
        return set(drives)

    def check_for_changes(self):
        """Check if USB drives have been added or removed"""
        current_drives = self.get_removable_drives()

        # Check for new drives
        new_drives = current_drives - self.previous_drives
        for drive in new_drives:
            self.usb_connected.emit(drive)

        # Check for removed drives
        removed_drives = self.previous_drives - current_drives
        for drive in removed_drives:
            self.usb_disconnected.emit(drive)

        self.previous_drives = current_drives


class USBFormatterWindow(QMainWindow):
    format_complete = pyqtSignal(str, bool, str)  # drive_letter, success, message
    progress_update = pyqtSignal(str, int, str)  # drive_letter, percentage, message

    def __init__(self):
        super().__init__()
        self.usb_monitor = USBMonitor()
        self.format_lock = Lock()
        self.drive_status = {}  # Track status of each drive
        self.drive_progress = {}  # Track progress percentage for each drive
        self.drive_sizes = {}  # Track drive sizes in GB
        self.drive_info = {}  # Track USB product information
        self.active_threads = {}  # Track active formatting threads by drive letter
        self.wmi_client = wmi.WMI()  # WMI client for hardware info
        self.init_ui()
        self.setup_monitoring()

    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("Incident Responder USB Formatter")
        self.setGeometry(100, 100, 900, 750)

        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Title
        title = QLabel("Incident Responder USB Formatter")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("color: #8B0000; padding: 10px;")  # Dark red color
        main_layout.addWidget(title)

        # Status group
        status_group = QGroupBox("Status")
        status_layout = QVBoxLayout()
        self.status_label = QLabel("Monitoring for USB devices...")
        self.status_label.setStyleSheet("color: blue; font-weight: bold;")
        status_layout.addWidget(self.status_label)

        # Active operations status
        self.active_label = QLabel("Active Operations: 0 drives formatting")
        self.active_label.setStyleSheet("color: gray;")
        status_layout.addWidget(self.active_label)

        status_group.setLayout(status_layout)
        main_layout.addWidget(status_group)

        # Connected Drives Table
        drives_group = QGroupBox("Connected USB Drives")
        drives_layout = QVBoxLayout()

        self.drives_table = QTableWidget()
        self.drives_table.setColumnCount(6)
        self.drives_table.setHorizontalHeaderLabels(["Drive", "Label", "Product Info", "Status", "Progress", "Actions"])

        # Make columns resizable by user
        header = self.drives_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)

        # Set default column widths
        self.drives_table.setColumnWidth(0, 60)   # Drive
        self.drives_table.setColumnWidth(1, 120)  # Label
        self.drives_table.setColumnWidth(2, 200)  # Product Info
        self.drives_table.setColumnWidth(3, 120)  # Status
        self.drives_table.setColumnWidth(4, 150)  # Progress
        self.drives_table.setColumnWidth(5, 100)  # Actions

        # Allow last section to stretch
        header.setStretchLastSection(True)

        # Enable word wrap for better text display
        self.drives_table.setWordWrap(True)

        # Set row height to auto-resize based on content
        self.drives_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)

        self.drives_table.setMinimumHeight(200)
        drives_layout.addWidget(self.drives_table)

        drives_group.setLayout(drives_layout)
        main_layout.addWidget(drives_group)

        # Drive selection group (for manual formatting)
        drive_group = QGroupBox("Manual Format")
        drive_layout = QHBoxLayout()
        drive_layout.addWidget(QLabel("Select Drive:"))
        self.drive_combo = QComboBox()
        drive_layout.addWidget(self.drive_combo)
        self.refresh_button = QPushButton("Refresh")
        self.refresh_button.clicked.connect(self.refresh_drives)
        drive_layout.addWidget(self.refresh_button)
        drive_group.setLayout(drive_layout)
        main_layout.addWidget(drive_group)

        # Format options group
        format_group = QGroupBox("Format Options")
        format_layout = QVBoxLayout()

        # File system selection
        fs_layout = QHBoxLayout()
        fs_layout.addWidget(QLabel("File System:"))
        self.filesystem_combo = QComboBox()
        self.filesystem_combo.addItems(["FAT32", "NTFS", "exFAT"])
        self.filesystem_combo.currentTextChanged.connect(self.on_filesystem_changed)
        fs_layout.addWidget(self.filesystem_combo)
        format_layout.addLayout(fs_layout)

        # Filesystem info/warning label
        self.fs_info_label = QLabel("")
        self.fs_info_label.setStyleSheet("color: #666; font-size: 9pt;")
        self.fs_info_label.setWordWrap(True)
        format_layout.addWidget(self.fs_info_label)

        # Secure erase option
        self.secure_erase_checkbox = QCheckBox("Secure Erase (Overwrite data before format)")
        self.secure_erase_checkbox.setStyleSheet("font-weight: bold; color: #8B0000;")
        self.secure_erase_checkbox.setToolTip(
            "Enable this option to securely overwrite all data on the drive before formatting.\n"
            "This makes data recovery extremely difficult. WARNING: This process will take significantly longer!"
        )
        format_layout.addWidget(self.secure_erase_checkbox)

        # Info label for secure erase
        secure_info = QLabel("⚠ Secure erase will overwrite the entire drive multiple times (much slower)")
        secure_info.setStyleSheet("color: #666; font-size: 9pt; font-style: italic;")
        secure_info.setWordWrap(True)
        format_layout.addWidget(secure_info)

        format_group.setLayout(format_layout)
        main_layout.addWidget(format_group)

        # Control buttons
        button_layout = QHBoxLayout()
        self.format_button = QPushButton("Format Selected Drive")
        self.format_button.setStyleSheet("background-color: #ff6b6b; color: white; font-weight: bold; padding: 10px;")
        self.format_button.clicked.connect(self.format_drive_manual)
        button_layout.addWidget(self.format_button)

        self.auto_format_button = QPushButton("Enable Auto-Format")
        self.auto_format_button.setCheckable(True)
        self.auto_format_button.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        self.auto_format_button.clicked.connect(self.toggle_auto_format)
        button_layout.addWidget(self.auto_format_button)

        # Debug button
        self.debug_button = QPushButton("Debug Scan")
        self.debug_button.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; padding: 10px;")
        self.debug_button.clicked.connect(self.run_debug_scan)
        button_layout.addWidget(self.debug_button)

        main_layout.addLayout(button_layout)

        # Log display
        log_group = QGroupBox("Activity Log")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(200)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)

        # Auto-format flag
        self.auto_format_enabled = False

        # Connect signals
        self.format_complete.connect(self.on_format_complete)
        self.progress_update.connect(self.on_progress_update)

        # Now refresh drives after all UI components are created
        self.refresh_drives()

        self.log("Application started. Monitoring for external drives...")
        self.log("Detecting: USB devices and portable SCSI drives (with surprise removal policy)")
        self.log("Multi-threaded support enabled - all drives will be formatted concurrently!")
        self.update_drives_table()

    def setup_monitoring(self):
        """Set up USB monitoring timer"""
        self.monitor_timer = QTimer()
        self.monitor_timer.timeout.connect(self.check_usb_changes)
        self.monitor_timer.start(1000)  # Check every second

        # Connect signals
        self.usb_monitor.usb_connected.connect(self.on_usb_connected)
        self.usb_monitor.usb_disconnected.connect(self.on_usb_disconnected)

    def check_usb_changes(self):
        """Check for USB drive changes"""
        self.usb_monitor.check_for_changes()

    def get_drive_size(self, drive_letter):
        """Get the size of a drive in GB"""
        try:
            drive_path = f"{drive_letter}:\\"
            free_bytes = ctypes.c_ulonglong(0)
            total_bytes = ctypes.c_ulonglong(0)

            ctypes.windll.kernel32.GetDiskFreeSpaceExW(
                ctypes.c_wchar_p(drive_path),
                None,
                ctypes.pointer(total_bytes),
                ctypes.pointer(free_bytes)
            )

            size_gb = total_bytes.value / (1024 ** 3)
            return round(size_gb, 2)
        except:
            return 0

    def get_usb_product_info(self, drive_letter):
        """Get USB product information including VID, PID, Serial Number"""
        try:
            # Get physical disk associated with the drive letter
            for physical_disk in self.wmi_client.Win32_DiskDrive():
                for partition in physical_disk.associators("Win32_DiskDriveToDiskPartition"):
                    for logical_disk in partition.associators("Win32_LogicalDiskToPartition"):
                        if logical_disk.DeviceID == f"{drive_letter}:":
                            # Extract USB information
                            model = physical_disk.Model or "Unknown"
                            serial = physical_disk.SerialNumber or "N/A"
                            interface_type = physical_disk.InterfaceType or "Unknown"

                            # Try to get PNP device ID for VID/PID
                            pnp_id = physical_disk.PNPDeviceID or ""
                            vid = "N/A"
                            pid = "N/A"

                            # Extract VID and PID from PNP Device ID
                            if "VID_" in pnp_id and "PID_" in pnp_id:
                                try:
                                    vid_start = pnp_id.index("VID_") + 4
                                    vid = pnp_id[vid_start:vid_start + 4]
                                    pid_start = pnp_id.index("PID_") + 4
                                    pid = pnp_id[pid_start:pid_start + 4]
                                except:
                                    pass

                            return {
                                'model': model,
                                'serial': serial,
                                'vid': vid,
                                'pid': pid,
                                'interface': interface_type,
                                'pnp_id': pnp_id
                            }

            return {
                'model': 'Unknown',
                'serial': 'N/A',
                'vid': 'N/A',
                'pid': 'N/A',
                'interface': 'Unknown',
                'pnp_id': 'N/A'
            }
        except Exception as e:
            return {
                'model': 'Error',
                'serial': 'N/A',
                'vid': 'N/A',
                'pid': 'N/A',
                'interface': 'Unknown',
                'pnp_id': str(e)
            }

    def on_usb_connected(self, drive_letter):
        """Handle USB connection event"""
        # Get drive size
        size_gb = self.get_drive_size(drive_letter)
        self.drive_sizes[drive_letter] = size_gb

        # Get USB product information
        product_info = self.get_usb_product_info(drive_letter)
        self.drive_info[drive_letter] = product_info

        # Log detailed information
        self.log(f"USB Drive detected: {drive_letter}:\\ ({size_gb} GB)")
        self.log(f"  Model: {product_info['model']}")
        self.log(f"  VID: {product_info['vid']} | PID: {product_info['pid']}")
        self.log(f"  Serial: {product_info['serial']}")

        # Add to drive status tracking
        self.drive_status[drive_letter] = "Connected"
        self.drive_progress[drive_letter] = 0

        # Update UI
        self.refresh_drives()
        self.update_drives_table()
        self.update_status_display()

        if self.auto_format_enabled:
            self.format_drive_auto(drive_letter)

    def on_usb_disconnected(self, drive_letter):
        """Handle USB disconnection event"""
        self.log(f"USB Drive removed: {drive_letter}:\\")

        # Remove from drive status tracking
        if drive_letter in self.drive_status:
            del self.drive_status[drive_letter]
        if drive_letter in self.drive_progress:
            del self.drive_progress[drive_letter]

        # Update UI
        self.refresh_drives()
        self.update_drives_table()
        self.update_status_display()

    def refresh_drives(self):
        """Refresh the list of available drives"""
        self.drive_combo.clear()
        drives = self.usb_monitor.get_removable_drives()
        if drives:
            for drive in sorted(drives):
                try:
                    volume_info = win32api.GetVolumeInformation(f"{drive}:\\")
                    label = volume_info[0] if volume_info[0] else "No Label"
                    size_gb = self.drive_sizes.get(drive, self.get_drive_size(drive))
                    self.drive_combo.addItem(f"{drive}:\\ ({label}) - {size_gb} GB", drive)
                except:
                    self.drive_combo.addItem(f"{drive}:\\", drive)
        else:
            self.drive_combo.addItem("No removable drives found")

        # Update filesystem info when drive changes
        self.drive_combo.currentIndexChanged.connect(self.update_filesystem_info)
        self.update_filesystem_info()

    def on_filesystem_changed(self, filesystem):
        """Handle filesystem selection change"""
        self.update_filesystem_info()

    def update_filesystem_info(self):
        """Update the filesystem information label"""
        if self.drive_combo.currentIndex() == -1 or self.drive_combo.currentText() == "No removable drives found":
            self.fs_info_label.setText("")
            return

        drive_letter = self.drive_combo.currentData()
        if not drive_letter:
            self.fs_info_label.setText("")
            return

        size_gb = self.drive_sizes.get(drive_letter, 0)
        filesystem = self.filesystem_combo.currentText()

        if filesystem == "FAT32":
            if size_gb > 32:
                self.fs_info_label.setText(f"⚠ WARNING: FAT32 max size is 32GB. Your drive is {size_gb} GB. Use NTFS or exFAT.")
                self.fs_info_label.setStyleSheet("color: red; font-size: 9pt; font-weight: bold;")
            else:
                self.fs_info_label.setText("✓ FAT32: Max compatibility (Windows, Mac, Linux)")
                self.fs_info_label.setStyleSheet("color: green; font-size: 9pt;")
        elif filesystem == "NTFS":
            self.fs_info_label.setText("✓ NTFS: Best for Windows, supports large files")
            self.fs_info_label.setStyleSheet("color: green; font-size: 9pt;")
        elif filesystem == "exFAT":
            self.fs_info_label.setText("✓ exFAT: Good compatibility, no size limits")
            self.fs_info_label.setStyleSheet("color: green; font-size: 9pt;")

    def toggle_auto_format(self):
        """Toggle auto-format mode"""
        self.auto_format_enabled = self.auto_format_button.isChecked()
        if self.auto_format_enabled:
            self.auto_format_button.setText("Disable Auto-Format")
            self.auto_format_button.setStyleSheet("background-color: #f44336; color: white; font-weight: bold; padding: 10px;")
            self.log("WARNING: Auto-format enabled! USB drives will be formatted automatically!")
            QMessageBox.warning(self, "Auto-Format Enabled",
                              "WARNING: All USB drives connected from now on will be automatically formatted!\n\n"
                              "ALL DATA WILL BE LOST!\n\n"
                              "Click 'Disable Auto-Format' to turn off this feature.")
        else:
            self.auto_format_button.setText("Enable Auto-Format")
            self.auto_format_button.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
            self.log("Auto-format disabled.")

    def validate_filesystem_for_drive(self, drive_letter, filesystem):
        """Validate if the filesystem is compatible with the drive size"""
        size_gb = self.drive_sizes.get(drive_letter, 0)

        # FAT32 has a 32GB limit in Windows
        if filesystem == "FAT32" and size_gb > 32:
            return False, f"FAT32 cannot be used on drives larger than 32GB.\n\nYour drive is {size_gb} GB.\n\nPlease select NTFS or exFAT instead."

        return True, ""

    def format_drive_manual(self):
        """Manually format the selected drive"""
        if self.drive_combo.currentIndex() == -1 or self.drive_combo.currentText() == "No removable drives found":
            QMessageBox.warning(self, "No Drive", "Please select a valid drive to format.")
            return

        drive_letter = self.drive_combo.currentData()
        filesystem = self.filesystem_combo.currentText()

        # Validate filesystem choice
        valid, error_msg = self.validate_filesystem_for_drive(drive_letter, filesystem)
        if not valid:
            QMessageBox.critical(self, "Invalid Filesystem", error_msg)
            return

        # Get drive size for display
        size_gb = self.drive_sizes.get(drive_letter, 0)

        # Build confirmation message
        secure_erase_enabled = self.secure_erase_checkbox.isChecked()
        secure_msg = "\n\nSECURE ERASE ENABLED: Drive will be securely overwritten before formatting.\nThis will take significantly longer!" if secure_erase_enabled else ""

        # Confirmation dialog
        reply = QMessageBox.question(self, "Confirm Format",
                                     f"Are you sure you want to format drive {drive_letter}:\\ ({size_gb} GB)?\n\n"
                                     f"ALL DATA ON THIS DRIVE WILL BE LOST!\n\n"
                                     f"File System: {filesystem}"
                                     f"{secure_msg}",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            self.format_drive(drive_letter)

    def format_drive_auto(self, drive_letter):
        """Automatically format a drive (called when auto-format is enabled)"""
        filesystem = self.filesystem_combo.currentText()

        # Validate filesystem choice
        valid, error_msg = self.validate_filesystem_for_drive(drive_letter, filesystem)
        if not valid:
            self.log(f"ERROR: Cannot format {drive_letter}:\\ - {error_msg.replace(chr(10), ' ')}")
            self.drive_status[drive_letter] = "Failed"
            self.update_drives_table()
            return

        self.log(f"Starting immediate format of drive {drive_letter}:\\...")
        self.start_format_thread(drive_letter)

    def start_format_thread(self, drive_letter):
        """Start a new thread to format the drive immediately"""
        # Check if already formatting this drive
        if drive_letter in self.active_threads and self.active_threads[drive_letter].is_alive():
            self.log(f"Drive {drive_letter}:\\ is already being formatted")
            return

        filesystem = self.filesystem_combo.currentText()
        secure_erase = self.secure_erase_checkbox.isChecked()

        erase_msg = " with secure erase" if secure_erase else ""
        self.log(f"Drive {drive_letter}:\\ formatting started as {filesystem}{erase_msg}")

        # Create and start a new thread for this drive
        format_thread = Thread(
            target=self.format_worker,
            args=(drive_letter, filesystem, secure_erase),
            daemon=True
        )
        self.active_threads[drive_letter] = format_thread
        format_thread.start()

        # Update status
        with self.format_lock:
            status_text = "Securely Erasing" if secure_erase else "Formatting"
            self.drive_status[drive_letter] = status_text
            self.drive_progress[drive_letter] = 0
            self.update_drives_table()
            self.update_active_status()

    def format_worker(self, drive_letter, filesystem, secure_erase):
        """Background worker that formats a single drive (runs in its own thread)"""
        try:
            # Log product information before formatting
            product_info = self.drive_info.get(drive_letter, {})
            if product_info:
                self.log(f"Formatting {drive_letter}:\\ - Product Info:")
                self.log(f"  VID:PID = {product_info.get('vid', 'N/A')}:{product_info.get('pid', 'N/A')}")
                self.log(f"  Serial: {product_info.get('serial', 'N/A')}")
                self.log(f"  Model: {product_info.get('model', 'Unknown')}")

            # Perform secure erase if requested
            if secure_erase:
                self.progress_update.emit(drive_letter, 10, "Starting secure erase...")
                success, message = self.perform_secure_erase(drive_letter)
                if not success:
                    self.format_complete.emit(drive_letter, False, f"Secure erase failed: {message}")
                    return

                self.progress_update.emit(drive_letter, 50, "Secure erase complete")

                # Update status to formatting after secure erase
                with self.format_lock:
                    self.drive_status[drive_letter] = "Formatting"
                    self.update_drives_table()

            else:
                self.progress_update.emit(drive_letter, 10, "Starting format...")

            # Perform the actual format
            self.progress_update.emit(drive_letter, 60 if secure_erase else 20, "Formatting drive...")
            success, message = self.perform_format(drive_letter, filesystem)

            if success:
                self.progress_update.emit(drive_letter, 100, "Complete!")

            # Signal completion
            self.format_complete.emit(drive_letter, success, message)

        except Exception as e:
            self.format_complete.emit(drive_letter, False, f"Unexpected error: {str(e)}")
        finally:
            # Clean up thread reference
            with self.format_lock:
                if drive_letter in self.active_threads:
                    del self.active_threads[drive_letter]
                self.update_active_status()

    def perform_secure_erase(self, drive_letter):
        """Perform secure erase by overwriting the drive (runs in background thread)"""
        drive_path = f"{drive_letter}:"

        try:
            self.log(f"Starting secure erase on {drive_path} - this may take a while...")

            # Method 1: Use cipher command for secure overwrite (Windows built-in)
            # /W removes data from unused disk space (effectively wipes free space)
            cipher_cmd = f'cipher /W:{drive_path}\\'

            self.log(f"Executing secure erase: {cipher_cmd}")

            # Start the process without waiting for it to complete
            process = subprocess.Popen(
                cipher_cmd,
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )

            # Monitor the process and update progress
            start_time = time.time()
            max_time = 1800  # 30 minutes timeout
            progress_step = 15  # Start at 15%
            last_progress = 10

            while True:
                # Check if process is still running
                poll_result = process.poll()

                if poll_result is not None:
                    # Process finished
                    stdout, stderr = process.communicate()

                    if poll_result == 0 or "successfully" in stdout.lower():
                        self.log(f"Secure erase completed on {drive_path}")
                        return True, f"Secure erase completed on {drive_path}"
                    else:
                        # Try alternative method: format with full format (not quick)
                        self.log(f"Cipher method unsuccessful, attempting full format...")
                        return self.perform_full_format_erase(drive_letter)

                # Check timeout
                elapsed = time.time() - start_time
                if elapsed > max_time:
                    process.kill()
                    return False, "Secure erase operation timed out (exceeded 30 minutes)"

                # Update progress gradually
                # Progress from 15% to 45% over time
                progress = min(45, 15 + int((elapsed / max_time) * 30))
                if progress > last_progress:
                    self.progress_update.emit(drive_letter, progress, f"Securely erasing... ({int(elapsed/60)} min)")
                    last_progress = progress

                # Wait a bit before checking again
                time.sleep(5)

        except Exception as e:
            return False, f"Secure erase error: {str(e)}"

    def perform_full_format_erase(self, drive_letter):
        """Perform a full format as an alternative secure erase method"""
        drive_path = f"{drive_letter}:"

        try:
            # Full format without /Q flag (writes zeros to entire disk)
            cmd = f'format {drive_path} /FS:NTFS /X /Y'

            self.log(f"Executing full format for secure erase: {cmd}")

            # Start the process
            process = subprocess.Popen(
                cmd,
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )

            # Monitor the process
            start_time = time.time()
            max_time = 1800  # 30 minutes timeout
            last_progress = 15

            while True:
                poll_result = process.poll()

                if poll_result is not None:
                    stdout, stderr = process.communicate()
                    if poll_result == 0:
                        return True, f"Full format (secure erase) completed on {drive_path}"
                    else:
                        error_msg = stderr if stderr else stdout
                        return False, f"Full format failed: {error_msg}"

                elapsed = time.time() - start_time
                if elapsed > max_time:
                    process.kill()
                    return False, "Full format operation timed out"

                # Update progress
                progress = min(45, 15 + int((elapsed / max_time) * 30))
                if progress > last_progress:
                    self.progress_update.emit(drive_letter, progress, f"Full format erasing... ({int(elapsed/60)} min)")
                    last_progress = progress

                time.sleep(5)

        except Exception as e:
            return False, f"Error: {str(e)}"

    def perform_format(self, drive_letter, filesystem):
        """Perform the actual formatting operation (runs in background thread)"""
        drive_path = f"{drive_letter}:"

        try:
            # Use Windows format command with quick format
            cmd = f'format {drive_path} /FS:{filesystem} /Q /X /Y'

            # Run the format command
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=60)

            if result.returncode == 0:
                return True, f"Successfully formatted {drive_path} as {filesystem}"
            else:
                error_msg = result.stderr if result.stderr else result.stdout
                return False, f"Format failed: {error_msg}"

        except subprocess.TimeoutExpired:
            return False, "Format operation timed out"
        except Exception as e:
            return False, f"Error: {str(e)}"

    def on_format_complete(self, drive_letter, success, message):
        """Handle format completion (runs in main thread)"""
        self.log(message)

        if success:
            self.drive_status[drive_letter] = "Formatted"
            self.drive_progress[drive_letter] = 100
        else:
            self.drive_status[drive_letter] = "Failed"
            self.drive_progress[drive_letter] = 0

        self.update_drives_table()
        self.update_active_status()
        self.update_status_display()

    def on_progress_update(self, drive_letter, percentage, message):
        """Handle progress updates (runs in main thread)"""
        self.drive_progress[drive_letter] = percentage
        self.log(f"{drive_letter}:\\ - {message} ({percentage}%)")
        self.update_drives_table()

    def format_drive(self, drive_letter):
        """Format the specified drive (manual or auto)"""
        self.start_format_thread(drive_letter)

    def update_drives_table(self):
        """Update the table showing all connected drives"""
        drives = self.usb_monitor.get_removable_drives()
        self.drives_table.setRowCount(len(drives))

        for idx, drive in enumerate(sorted(drives)):
            # Drive letter
            self.drives_table.setItem(idx, 0, QTableWidgetItem(f"{drive}:\\"))

            # Volume label
            try:
                volume_info = win32api.GetVolumeInformation(f"{drive}:\\")
                label = volume_info[0] if volume_info[0] else "No Label"
            except:
                label = "Unknown"
            self.drives_table.setItem(idx, 1, QTableWidgetItem(label))

            # Product Information
            product_info = self.drive_info.get(drive, {})
            if product_info:
                vid = product_info.get('vid', 'N/A')
                pid = product_info.get('pid', 'N/A')
                serial = product_info.get('serial', 'N/A')
                model = product_info.get('model', 'Unknown')
                product_text = f"VID:{vid} PID:{pid}\nS/N:{serial}\nModel:{model}"
            else:
                product_text = "Loading..."

            product_item = QTableWidgetItem(product_text)
            # Set text alignment and enable wrapping
            product_item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            self.drives_table.setItem(idx, 2, product_item)

            # Status
            status = self.drive_status.get(drive, "Connected")
            status_item = QTableWidgetItem(status)

            # Color code based on status
            if status == "Formatting":
                status_item.setBackground(QColor("#FFA500"))  # Orange
            elif status == "Securely Erasing":
                status_item.setBackground(QColor("#FF8C00"))  # Dark orange
                status_item.setForeground(QColor("#FFFFFF"))  # White text
            elif status == "Formatted":
                status_item.setBackground(QColor("#90EE90"))  # Light green
            elif status == "Failed":
                status_item.setBackground(QColor("#FFB6C1"))  # Light red

            self.drives_table.setItem(idx, 3, status_item)

            # Progress bar
            progress = self.drive_progress.get(drive, 0)
            progress_bar = QProgressBar()
            progress_bar.setValue(progress)
            progress_bar.setMinimum(0)
            progress_bar.setMaximum(100)

            # Style progress bar based on status
            if status == "Securely Erasing":
                progress_bar.setStyleSheet("""
                    QProgressBar {
                        border: 2px solid #999;
                        border-radius: 5px;
                        text-align: center;
                        background-color: #f0f0f0;
                    }
                    QProgressBar::chunk {
                        background-color: #FF8C00;
                    }
                """)
            elif status == "Formatting":
                progress_bar.setStyleSheet("""
                    QProgressBar {
                        border: 2px solid #999;
                        border-radius: 5px;
                        text-align: center;
                        background-color: #f0f0f0;
                    }
                    QProgressBar::chunk {
                        background-color: #FFA500;
                    }
                """)
            elif status == "Formatted":
                progress_bar.setStyleSheet("""
                    QProgressBar {
                        border: 2px solid #999;
                        border-radius: 5px;
                        text-align: center;
                        background-color: #f0f0f0;
                    }
                    QProgressBar::chunk {
                        background-color: #4CAF50;
                    }
                """)
            elif status == "Failed":
                progress_bar.setStyleSheet("""
                    QProgressBar {
                        border: 2px solid #999;
                        border-radius: 5px;
                        text-align: center;
                        background-color: #f0f0f0;
                    }
                    QProgressBar::chunk {
                        background-color: #f44336;
                    }
                """)

            self.drives_table.setCellWidget(idx, 4, progress_bar)

            # Action button
            action_text = "Format" if status not in ["Formatting", "Securely Erasing"] else "In Progress"
            self.drives_table.setItem(idx, 5, QTableWidgetItem(action_text))

    def update_active_status(self):
        """Update the active operations status label"""
        # Count active threads
        active_count = sum(1 for thread in self.active_threads.values() if thread.is_alive())
        self.active_label.setText(f"Active Operations: {active_count} drive(s) formatting")

        if active_count > 0:
            self.active_label.setStyleSheet("color: orange; font-weight: bold;")
        else:
            self.active_label.setStyleSheet("color: gray;")

    def update_status_display(self):
        """Update the main status display"""
        drives = self.usb_monitor.get_removable_drives()
        num_drives = len(drives)

        if num_drives == 0:
            self.status_label.setText("Monitoring for USB devices...")
            self.status_label.setStyleSheet("color: blue; font-weight: bold;")
        elif num_drives == 1:
            drive = list(drives)[0]
            status = self.drive_status.get(drive, "Connected")
            self.status_label.setText(f"1 USB drive connected ({drive}:\\) - Status: {status}")
            self.status_label.setStyleSheet("color: green; font-weight: bold;")
        else:
            self.status_label.setText(f"{num_drives} USB drives connected")
            self.status_label.setStyleSheet("color: green; font-weight: bold;")

    def run_debug_scan(self):
        """Run a debug scan of all drives and log details"""
        self.log("=" * 50)
        self.log("DEBUG SCAN: Scanning all drives...")
        self.log("=" * 50)

        # Redirect stdout to capture debug messages
        import io
        from contextlib import redirect_stdout

        debug_output = io.StringIO()
        with redirect_stdout(debug_output):
            drives = self.usb_monitor.get_removable_drives(debug=True)

        # Log all debug output
        for line in debug_output.getvalue().split('\n'):
            if line.strip():
                self.log(line)

        self.log("=" * 50)
        self.log(f"DEBUG SCAN COMPLETE: Found {len(drives)} external drive(s): {sorted(drives)}")
        self.log("=" * 50)

    def log(self, message):
        """Add a message to the log"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")


def main():
    app = QApplication(sys.argv)

    # Check if running as administrator on Windows
    try:
        import ctypes
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
        if not is_admin:
            QMessageBox.warning(None, "Administrator Rights Required",
                              "This application requires Administrator privileges to format drives.\n\n"
                              "Please right-click the script and select 'Run as Administrator'.")
    except:
        pass

    window = USBFormatterWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
