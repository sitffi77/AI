import sys
from PyQt5.QtWidgets import QApplication
from ui import PainDiagnosisApp


def main():
    print("🚀 Запуск приложения диагностики боли...")
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    window = PainDiagnosisApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()