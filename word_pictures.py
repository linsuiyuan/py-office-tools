import os, os.path as path
from docx import Document, ImagePart
from docx.image.image import Image
from docx.opc.packuri import PackURI


def extract_pics_from_docx(filename, image_dir="imgs"):
    """
    从 docx 文件中提取图片
        :param filename: 待提取文件名称
        :param image_dir: 存放提取图片的目录
        :return:
    """
    ext = path.splitext(filename)[-1]
    if ext != '.docx':
        raise Exception("只支持提取 docx文件 的图片")

    # 如果是相对路径转为绝对路径
    if not path.isabs(image_dir):
        image_dir = path.join(path.dirname(filename), image_dir)
    if not path.exists(image_dir):
        os.mkdir(image_dir)

    doc = Document(filename)

    image_parts = doc.part.package.image_parts._image_parts
    for part in image_parts:
        p_filename = path.join(image_dir, path.basename(part.partname))
        with open(p_filename, 'wb') as f:
            f.write(part.blob)
    print("图片导出完成!")


def replace_pic(doc, from_pic, to_pic):
    """
    图片替换
        :param doc: docx.Document 对象
        :param from_pic:
        :param to_pic:
        :return:
    """
    related_parts = doc.part.related_parts
    parts = doc.part.package.parts

    # 替换 part
    def replace_part(items, raw_items):
        for k, p in items:
            if path.basename(p.partname) == from_pic:
                image = Image.from_file(to_pic)
                partname = path.join(path.dirname(p.partname), image.filename)
                partname = PackURI(partname)
                img_part = ImagePart.from_image(image, partname)
                raw_items.__setitem__(k, img_part)
                break
    replace_part(related_parts.items(), related_parts)
    replace_part(enumerate(parts), parts)

    # 同步替换的 part 到 rels
    rels = doc.part.rels
    for rid, p in rels.related_parts.items():
        if rels[rid].target_part is not p:
            rels[rid]._target = p
            break

    return doc


def replace_pics(pic_pairs, input_file, output_file):
    """
    批量图片替换
        :param pic_pairs:
        :param input_file:
        :param output_file:
        :return:
    """
    doc = Document(input_file)

    # 如果 pic_pairs里的替换图片是相对路径, 则转为绝对路径, 使用相对路径时, 图片应放在和 word文件相同的目录下
    pic_pairs = [(f, t) if path.isabs(t)
                 else (f, path.join(path.dirname(input_file), t))
                 for f, t in pic_pairs]

    [replace_pic(doc, f, t) for f, t in pic_pairs]

    doc.save(output_file)


if __name__ == '__main__':
    input_file = "包罗万有.docx"
    output_file = "包罗万有1.docx"
    pic_pairs = [
        ('image5.jpeg', 'WechatIMG49_1.png'),
        ('image6.jpeg', 'WechatIMG42_1.png')
    ]
    replace_pics(pic_pairs, input_file, output_file)
