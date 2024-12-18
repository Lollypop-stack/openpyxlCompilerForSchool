import logging
from kundolukParser.gui import GUI

def main():
    session = "l6tsri25mpjhraetv88nlvhnm206ruvl"
    gui = GUI(session)
    gui.start()

if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,  # Уровень логирования (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"  # Формат сообщения
    )

    main()
