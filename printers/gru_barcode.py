from barcode import Code39, Code128, writer
from barcode.writer import ImageWriter
import qrcode

from app_config import BARCODE_FOLDER_ROOT_PATH


writer_options = {
    'default': {
        'font_size': 15,
        'text_distance': 1.0,
        'dpi': 300,
        'module_height': 15.0,
    },
    'user_id': {
        'font_size': 6,
        'text_distance': 1.0,
        'dpi': 160,
        'module_height': 6.0,
    },
    'shipping_mark': {
        'font_size': 10,
        'text_distance': 1.0,
        'dpi': 300,
        'module_height': 20.0,
    }
}


def create_barcode(case_no, code, label_type='default', file_path=BARCODE_FOLDER_ROOT_PATH):
    file_name = '{0}_{1}.{2}'.format(case_no, code, 'jpg')
    barcode_image_path = file_path.format(file_name)

    with open(file=barcode_image_path, mode='wb') as f:
        if code == 'code39':
            t = Code39(case_no, writer=ImageWriter(), add_checksum=False)
            t.write(f, options=writer_options[label_type])

        if code == 'code128':
            t = Code128(case_no, writer=ImageWriter())
            t.write(f, options=writer_options[label_type])

        if code == 'qr':
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_H,
                box_size=4,
                border=4,
            )
            qr.add_data(case_no)
            qr.make(fit=True)
            img = qr.make_image()
            img.save(barcode_image_path)
    return barcode_image_path


if __name__ == "__main__":
    # create_barcode('SLM071109RZ1711A', 'code39')
    # create_barcode('SLM071109RZ1711A', 'code128')
    # create_barcode('SLM071109RZ1711A', 'qr')
    pass
