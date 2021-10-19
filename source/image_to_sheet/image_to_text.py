import os
import shutil
import set_credentials    # separate file


def detect_text(path):
    """Detects text in the file.
    Copied from Google Cloud vision docs"""
    from google.cloud import vision
    import io
    client = vision.ImageAnnotatorClient()

    with io.open(path, 'rb') as image_file:
        content = image_file.read()

    image = vision.Image(content=content)

    response = client.text_detection(image=image)
    text = response.text_annotations

    if response.error.message:
        raise Exception(
            '{}\nFor more info on error messages, check: '
            'https://cloud.google.com/apis/design/errors'.format(
                response.error.message))
    return text


def save_text(path, text):
    with open(path, 'w', encoding="utf-8") as file:
        file.write(text.description)


def images_to_texts():
    images = os.listdir('images/unread')
    if len(images) > 0:
        for image in images:
            image_name = image.split('.')[0]
            text = detect_text(path='images/unread/' + image)
            save_text('texts/unread/' + image_name  + '.txt', text[0])
            shutil.move('images/unread/' + image, 'images/read/' + image)
            print('Finished reading ' + image)


if __name__ == '__main__':
    set_credentials.credential_path()
    os.chdir('../..')
    images_to_texts()


