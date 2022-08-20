from pptx.util import Inches


class Model:
    # get ratio for both and use the smaller one to ensure that the image would fit in the slide panel
    def get_resize_ratio(self, img_width, img_height, panel_pixel_width, panel_pixel_height):
        ratio_width = panel_pixel_width / img_width
        ratio_height = panel_pixel_height / img_height
        return min(ratio_width, ratio_height)

    def get_margin(self, length, resized_length, dpi):
        return Inches(length - resized_length / dpi) / 2

    def get_margin_in_pixel(self,length,resized_length):
        return (length-resized_length)/2

    def check_image_extension(self,lst,file):
        if file.endswith('.png') or file.endswith(".tif") or file.endswith(".jpg") or file.endswith(".jpeg") or file.endswith(".PNG"):
            lst.append(file)
            return lst

    def apply_vertical_margin(self, index, row, column, prefixed_vertical_position, margin_height):
        # apply margin
        iter = row * column
        if 0 <= index % iter < column:
            fixed_vertical_position = prefixed_vertical_position
        elif column * (row - 1) <= index % iter < iter:
            fixed_vertical_position = prefixed_vertical_position + margin_height * 2
        else:
            fixed_vertical_position = prefixed_vertical_position + margin_height
        return fixed_vertical_position

