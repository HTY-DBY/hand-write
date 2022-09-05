import docx
import os
import shutil
from PIL import Image, ImageFont
from handright import Template, handwrite
from multiprocessing import Pool
# 帮助文档 https://github.com/Gsllchb/Handright/blob/master/docs/tutorial.md


def get_text(file_path, indent_size=6):
    # 自动缩进排版，如果已在word里设置缩进可以注释本段
    # file_path文档路径，indent_size控制缩进
    doc = docx.Document('text/' + file_path)
    texts = []
    indent = ''
    for i in range(0, indent_size):
        indent = indent + ' '
    for paragraph in doc.paragraphs:
        texts.append(indent + paragraph.text)
    return '\n'.join(texts)


def rea(path, pdf_name):
    file_list = os.listdir(path)
    pic_name = []
    im_list = []
    for x in file_list:
        if "jpg" in x or 'png' in x or 'jpeg' in x:
            pic_name.append(x)

    pic_name.sort()
    new_pic = []

    for x in pic_name:
        if "jpg" in x:
            new_pic.append(x)

    for x in pic_name:
        if "png" in x:
            new_pic.append(x)

    # print("hec", new_pic)

    im1 = Image.open(os.path.join(path, new_pic[0]))
    new_pic.pop(0)
    for i in new_pic:
        img = Image.open(os.path.join(path, i))
        # im_list.append(Image.open(i))
        if img.mode == "RGBA":
            img = img.convert('RGB')
            im_list.append(img)
        else:
            im_list.append(img)
    im1.save(pdf_name, "PDF", resolution=100.0,
             save_all=True, append_images=im_list)


if __name__ == "__main__":
    # 重置 save 文件夹
    shutil.rmtree('save')
    os.mkdir('save')

    text = get_text('temp.docx')  # 根目录下的word文档
    f_type = '01.ttf'  # 字体选择
    hty_bj = 100  # 页边距
    hty_size = 80  # 字体大小

    template = Template(
        # background=Image.open('bac/01.png'),  # 自定义背景图片
        # 这是 300ppi 的标准 A4
        background=Image.new(mode="1", size=(2479, 3508), color=1),
        font=ImageFont.truetype("font/"+f_type, size=hty_size),  # 手写字体选择
        line_spacing=hty_size+15,
        fill=0,  # 字体颜色
        left_margin=hty_bj,
        top_margin=hty_bj,
        right_margin=hty_bj,
        bottom_margin=hty_bj,
        # word_spacing=5,
        line_spacing_sigma=3,  # 行间距随机扰动
        font_size_sigma=4,  # 字体大小随机扰动
        word_spacing_sigma=2,  # 字间距随机扰动
        end_chars="，。",  # 防止特定字符因排版算法的自动换行而出现在行首
        # perturb_x_sigma=2,  # 笔画横向偏移随机扰动
        # perturb_y_sigma=2,  # 笔画纵向偏移随机扰动
        # perturb_theta_sigma=0.05,  # 笔画旋转偏移随机扰动
    )

    print('----请等待----')
    with Pool() as p:
        # 并行计算，这一步开向量池，较慢
        images = handwrite(text, template, mapper=p.map)
    print('----正在生成----')
    for i, im in enumerate(images):
        assert isinstance(im, Image.Image)
        # im.show()
        im.save("save/{}.jpg".format(i))
        print('----生成----'+format(i)+'图')
    print('----已生成所有, 合成pdf中----')

    pdf_name = 'result.pdf'
    mypath = r"save"
    if ".pdf" in pdf_name:
        rea(mypath, pdf_name=pdf_name)
    else:
        rea(mypath, pdf_name="{}.pdf".format(pdf_name))
    print('----已合成pdf----')

    os.popen(r'result.pdf')
