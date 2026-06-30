import sys
import os
import subprocess
import importlib
from pathlib import Path


def _ensure_dependencies():
    required_packages = {
        "PyQt5": "PyQt5",
        "pandas": "pandas",
        "numpy": "numpy",
        "matplotlib": "matplotlib",
        "openpyxl": "openpyxl",
        "xlrd": "xlrd",
    }

    for module_name, pip_name in required_packages.items():
        try:
            importlib.import_module(module_name)
        except ImportError:
            print(f"Eksik paket bulundu: {pip_name}. Kurulum başlatılıyor...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])
            except Exception as exc:
                print(f"{pip_name} kurulamadı: {exc}")
                print("Lütfen şu komutu çalıştırın:")
                print(f"  {sys.executable} -m pip install {pip_name}")
                raise


_ensure_dependencies()


def _configure_qt_environment():
    """
    Qt platform/plugin ayarlarını çalışma ortamına göre güvenli hale getirir.
    Özellikle Windows dışında `QT_QPA_PLATFORM=windows` kaldığında uygulama açılmaz.
    """
    current_platform = os.environ.get("QT_QPA_PLATFORM", "").strip().lower()
    if current_platform == "windows" and sys.platform != "win32":
        # Linux/macOS üzerinde yanlış "windows" platformu atanmışsa temizle.
        os.environ.pop("QT_QPA_PLATFORM", None)

    # PyQt5 platform plugin dizinini gerçek paket konumundan bul.
    # Windows'ta PyQt5 kullanıcı dizinine kurulmuş olabilir; bu durumda sys.executable
    # altındaki site-packages varsayımı qwindows.dll yolunu bulamaz.
    import PyQt5

    pyqt_package_dir = Path(PyQt5.__file__).resolve().parent
    executable_dir = Path(sys.executable).resolve().parent
    version = f"python{sys.version_info.major}.{sys.version_info.minor}"
    candidates = [
        pyqt_package_dir / "Qt5" / "plugins" / "platforms",
        pyqt_package_dir / "Qt" / "plugins" / "platforms",
        pyqt_package_dir / "plugins" / "platforms",
        executable_dir / "Lib" / "site-packages" / "PyQt5" / "Qt5" / "plugins" / "platforms",
        executable_dir / "lib" / version / "site-packages" / "PyQt5" / "Qt5" / "plugins" / "platforms",
    ]

    for platform_dir in candidates:
        if platform_dir.exists():
            os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = str(platform_dir)
            plugin_root = platform_dir.parent
            existing_plugin_path = os.environ.get("QT_PLUGIN_PATH", "")
            if str(plugin_root) not in existing_plugin_path.split(os.pathsep):
                os.environ["QT_PLUGIN_PATH"] = (
                    str(plugin_root)
                    if not existing_plugin_path
                    else str(plugin_root) + os.pathsep + existing_plugin_path
                )
            break


_configure_qt_environment()

from PyQt5.QtWidgets import QApplication

from spul.spul_app import SledAnalyzerApp


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SledAnalyzerApp()
    window.show()
    sys.exit(app.exec_())
