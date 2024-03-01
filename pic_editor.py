from PIL import Image
import os


def compress_images(input_folder, output_folder, target_size=(270, 240)):
    # Проверка, что выходная папка существует или создать ее
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

        # Получаем список файлов входной папки
    input_files = os.listdir(input_folder)
    for input_file in input_files:
        input_path = os.path.join(input_folder, input_file)
        output_path = os.path.join(output_folder, input_file)
        # Открываем изображение
        img = Image.open(input_path)
        # Масштабируем изображение до целевого размера
        img.thumbnail(target_size)
        # Сохраняем сжатое изображение
        img.save(output_path)


if __name__ == "__main__":
    # Укажите путь к папке с исходными изображениями и папке для сохранения сжатых изображений
    input_folder_ = r"C:\Users\Administrator\Desktop\Изображения для подкатегорий\До обработки"
    output_folder_ = r"C:\Users\Administrator\Desktop\Изображения для подкатегорий\После обработки"

    # Укажите целевой размер в пикселях
    target_size_ = (270, 240)

    # Вызываем функцию для сжатия изображений
    compress_images(input_folder_, output_folder_, target_size_)
