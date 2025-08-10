---
title: 使用LLM 实现一个可编辑PPT Agent
date: 2025-08-10 
categories: [LLM, Agent]
tags: [agent, 工作流]     # TAG 名称应始终小写
author: LiuKe
---

# 目的与实现思路

最近，很多AI厂商开始提供深度调研后，将调研内容整理为网页的功能，如谷歌，秘塔都有类似的服务。使用时发现生成的网页相对精美，内容并不比AIPPT这类专做PPT生成工具的网站差。考虑到ppt的本质就是open xml，而且ppt格式给出（无前端基础的）用户一定的自由度，可以自己调整页面内容，排版等，因此，我认为将llm生成html的能力与ppt生成结合起来，会是一个很不错的工具。 

具体的实现思路分为两版， 一版是直接使用llm+function call生成open xml的所有文档，最后打包修改为pptx格式，但尝试过程中发现，我对openxml的不熟悉极大地影响了实现效率，而能够一开始就生成可以打开的pptx文件的llm几乎不存在。

退而求其次，使用python pptx这类库同样可以实现一定的功能，但相对上一种实现做html2openxml的翻译，这种实现复杂的地方在于将html与ppt的操作对应起来，目前我只实现了一个较为粗糙的版本（在claude的协助下）。

# Prompt

完整的prompt如下。 
```
你是一位专业的网页前端设计师，精通 HTML 和 CSS。你的任务是为一次演示文稿（PPT）创建多页 HTML 内容。




---



## 任务和规则



1. **画布尺寸**：整个演示文稿的基准尺寸是 `1920x1080` 像素。

2. **页面容器**：每一页幻灯片都必须用一个 `<div>` 容器包裹，并且该容器必须拥有类名 `class="slide-page"`。

3. **绝对定位**：容器内的所有元素（文本、图片、形状）都必须使用 **内联 CSS** 并设定 `position: absolute`，同时要精确指定 `top`, `left`, `width`, `height` 这四个属性，单位为 `px`。

4. **文本元素**：

* 标题使用 `<h1>` 或 `<h2>` 标签，段落使用 `<p>` 标签。

* 所有文本元素都必须有 `font-size: ...px;`、`color: #RRGGBB;` 和 `font-family: ...;` 样式。

* **字体**：请使用"Microsoft YaHei"。

* **对齐**：使用 `text-align: center;` 或 `text-align: left;` 等样式来控制文本对齐。

* **加粗**：使用 `font-weight: bold;` 来加粗文本。
4. **元素定位**：

* 容器内的所有元素都必须使用 **内联 CSS** 并设定 `position: absolute`，同时精确指定 `top`, `left`, `width`, `height` 这四个属性，单位为 `px`。

* **宽度和换行**：

* **文本元素（`<h1>`, `<p>`等）的 `width` 属性必须足够宽，以容纳其文本内容，避免出现单字符宽度或过度拥挤。请确保宽度值大于文本的视觉长度。**

* 如果你希望文本在某个特定位置换行，请直接在文本中插入 `<br>` 标签。

* **居中规则**：如果需要水平居中，请同时使用 `left: 50%;` 和 `transform: translateX(-50%);`。
5. **形状元素**：

* 用一个 `<div>` 标签来表示带有背景色的形状。

* 形状的颜色必须用 `background-color: #RRGGBB;` 样式指定。


6. **图片元素**：

* 图片使用 `<img>` 标签，并提供有效的 `src` 属性（可以是图片 URL 或 Base64 数据）。

7. **内容要求**：

* 请生成三页幻灯片，主题是“人工智能在教育领域的应用”。

* 第一页是封面，包含标题和副标题。

* 第二页是核心内容，讨论个性化学习。

* 第三页是总结或展望。



---



请严格遵循上述规则，生成符合要求的 HTML 代码。请只输出 HTML 代码，不要包含任何额外的解释性文字。

```
以下是使用gpt-40 生成的网页。
```
Copilot said: <div class="slide-page" style="width: 1920px; height: 1080px; position:
<div class="slide-page" style="width: 1920px; height: 1080px; position: relative;"> <h1 style="position: absolute; top: 120px; left: 50%; transform: translateX(-50%); width: 1500px; font-size: 96px; color: #1b263b; text-align: center; font-weight: bold; font-family: 'Microsoft YaHei';">MCP 模块化组件平台介绍</h1> <p style="position: absolute; top: 280px; left: 50%; transform: translateX(-50%); width: 1200px; font-size: 48px; color: #324e7b; text-align: center; font-family: 'Microsoft YaHei'; font-weight: bold;">什么是 MCP？</p> <p style="position: absolute; top: 360px; left: 360px; width: 1200px; font-size: 38px; color: #22223b; text-align: left; font-family: 'Microsoft YaHei';"> MCP（Modular Component Platform）是一种面向企业应用的模块化开发与集成平台。通过组件化、可配置和高扩展性，MCP 能够帮助团队高效搭建复杂系统，实现业务快速迭代和灵活定制，提升研发效率和系统稳定性。 </p> <div style="position: absolute; top: 540px; left: 300px; width: 1320px; height: 3px; background-color: #8ecae6;"></div> <p style="position: absolute; top: 570px; left: 50%; transform: translateX(-50%); width: 1100px; font-size: 44px; color: #0077b6; text-align: center; font-family: 'Microsoft YaHei'; font-weight: bold;">MCP 的主要使用方式</p> <p style="position: absolute; top: 660px; left: 370px; width: 1100px; font-size: 34px; color: #293241; text-align: left; font-family: 'Microsoft YaHei';"> 1. 通过平台提供的可视化界面拖拽组装业务模块，快速搭建应用。<br> 2. 利用内置组件库，灵活复用已有功能模块，减少重复开发。<br> 3. 支持自定义组件开发，满足个性化业务需求。<br> 4. 通过统一接口和协议，实现各模块的无缝集成和数据互通。<br> 5. 提供完善的权限管理和监控能力，保障系统安全与稳定。 </p> <img src="https://img.alicdn.com/imgextra/i3/O1CN01Gf9QbF1y7yQYp5Xxw_!!6000000006549-2-tps-800-400.png" style="position: absolute; top: 830px; left: 50%; transform: translateX(-50%); width: 750px; height: 180px; object-fit: contain; border-radius: 20px;"> </div>
``` 


# function
有了生成的html之后，就应该考虑如何将html转为ppt，以下是代码。 


```

import os
import re
import base64
import requests
from typing import Dict, Optional, Tuple
from bs4 import BeautifulSoup, Tag
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from io import BytesIO

class LLMToPPTXConverter:
    """
    一个将LLM生成的HTML内容转换为PPTX的转换器。
    支持递归解析所有绝对定位的元素，累加父容器的偏移量。
    """

    def __init__(self, output_path: str, html_base_width: int = 1920, html_base_height: int = 1080):
        self.output_path = output_path
        self.prs = Presentation()
        EMU_PER_PX = 9525
        self.prs.slide_width = Emu(html_base_width * EMU_PER_PX)
        self.prs.slide_height = Emu(html_base_height * EMU_PER_PX)
        self.blank_layout = self.prs.slide_layouts[6]
        print(f"PPT页面尺寸已设置为: {html_base_width}x{html_base_height} px")

    def _parse_color(self, color_str: str) -> Optional[Tuple[int, int, int]]:
        if not color_str or color_str.lower() in ('transparent', 'inherit', 'none'):
            return None
        color_str = color_str.strip().lower()
        if color_str.startswith('#'):
            hex_color = color_str[1:]
            if len(hex_color) == 3:
                hex_color = ''.join([c * 2 for c in hex_color])
            if len(hex_color) == 6:
                try:
                    return (int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))
                except ValueError:
                    pass
        rgb_match = re.match(r'rgba?\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)', color_str)
        if rgb_match:
            return tuple(map(int, rgb_match.groups()))
        return None

    def _parse_size_to_emu(self, size_str: str, default_value: int = 0) -> Emu:
        size_str = size_str.strip().lower()
        pixel_value = default_value
        if size_str.endswith('px'):
            try:
                pixel_value = float(re.search(r'[\d.]+', size_str).group())
            except (ValueError, AttributeError):
                pass
        return Emu(pixel_value * 9525)
    
    def _px_to_pt(self, px: float) -> Pt:
        return Pt(px * 0.75)

    def _get_element_styles(self, element: Tag) -> Dict:
        styles = {}
        if element.get('style'):
            for declaration in element.get('style').split(';'):
                if ':' in declaration:
                    prop, value = declaration.split(':', 1)
                    styles[prop.strip()] = value.strip()
        return styles

    def _download_image(self, url: str) -> Optional[BytesIO]:
        try:
            if url.startswith('data:image'):
                header, data = url.split(',', 1)
                image_bytes = base64.b64decode(data)
                if image_bytes:
                    return BytesIO(image_bytes)
                else:
                    print(f"警告: Base64图片数据为空。")
                    return None
            response = requests.get(url, timeout=10, headers={'User-Agent': 'Mozilla/5.0'})
            response.raise_for_status()
            content_type = response.headers.get('content-type', '')
            if not content_type.startswith('image/'):
                print(f"警告: URL返回的内容不是图片。Content-Type: {content_type}")
                return None
            image_data = BytesIO(response.content)
            try:
                from PIL import Image
                img = Image.open(image_data)
                img.verify()
                image_data.seek(0)
                return image_data
            except Exception as img_e:
                print(f"警告: 下载的图片文件损坏或无效。Pillow验证失败: {img_e}")
                return None
        except requests.exceptions.RequestException as req_e:
            print(f"下载图片失败，请求异常：{req_e}")
            return None
        except Exception as e:
            print(f"下载图片失败，发生未知错误：{e}")
            return None

    def _to_px(self, size_str):
        # 工具函数，将 '120px' -> 120
        if isinstance(size_str, (int, float)):
            return float(size_str)
        if not size_str:
            return 0
        size_str = size_str.strip().lower()
        if size_str.endswith('px'):
            try:
                return float(re.search(r'[\d.]+', size_str).group())
            except Exception:
                return 0
        return 0

    def _get_text_with_br(self, element):
        """
        将<p>等标签中的<br>转换为换行符，并保留原始HTML中的物理换行。
        """
        text = ""
        for child in element.children:
            if isinstance(child, Tag) and child.name == "br":
                text += "\n"
            elif isinstance(child, Tag):
                text += self._get_text_with_br(child)
            elif child:
                # 保留物理换行
                text += str(child).replace('\r\n', '\n').replace('\r', '\n')
        return text
    def _add_element_to_slide_with_abs(self, slide, element, left_px, top_px):
        """
        类似 _add_element_to_slide，但接受绝对 px 坐标。
        """
        styles = self._get_element_styles(element)
        width = self._to_px(styles.get('width', '100px'))
        height = self._to_px(styles.get('height', '50px'))
        x_pos = self._parse_size_to_emu(f'{left_px}px')
        y_pos = self._parse_size_to_emu(f'{top_px}px')
        width_emu = self._parse_size_to_emu(f'{width}px')
        height_emu = self._parse_size_to_emu(f'{height}px')

        # 检查文本宽度
        if element.get_text(strip=True):
            min_width_px = 100
            if width < min_width_px:
                width = min_width_px
                width_emu = self._parse_size_to_emu(f'{min_width_px}px')

        if element.name == 'img':
            img_src = element.get('src')
            if img_src:
                image_data = self._download_image(img_src)
                if image_data:
                    slide.shapes.add_picture(image_data, x_pos, y_pos, width_emu, height_emu)
        elif element.name == 'div' and 'background-color' in styles:
            bg_color_str = styles.get('background-color')
            bg_color = self._parse_color(bg_color_str)
            if bg_color:
                is_circle = 'border-radius' in styles and styles['border-radius'] == '50%'
                shape_enum = MSO_SHAPE.RECTANGLE
                if is_circle and width == height:
                    shape_enum = MSO_SHAPE.OVAL
                shape = slide.shapes.add_shape(shape_enum, x_pos, y_pos, width_emu, height_emu)
                fill = shape.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(*bg_color)
                shape.line.fill.background()
            transform_style = styles.get('transform')
            if transform_style:
                rotate_match = re.search(r'rotate\(([-]?\d+)(deg)?\)', transform_style)
                if rotate_match:
                    try:
                        rotation_value = float(rotate_match.group(1))
                        shape.rotation = rotation_value
                        print(f"应用旋转样式: {rotation_value}度")
                    except (ValueError, IndexError):
                        print("警告: 无法解析旋转角度。")
        elif element.get_text(strip=True):
            # 处理文本和换行
            text_content = self._get_text_with_br(element) if element.find("br") else element.get_text(strip=True)
            font_size_px = float(re.search(r'[\d.]+', styles.get('font-size', '16px')).group())
            font_size_pt = self._px_to_pt(font_size_px)
            width_emu = self._parse_size_to_emu(styles.get('width', '400px'))
            height_emu = self._parse_size_to_emu(styles.get('height', '1000px'))  # 你可以设置为容器高度或页面高度

            textbox = slide.shapes.add_textbox(x_pos, y_pos, width_emu, height_emu)
            text_frame = textbox.text_frame
            text_frame.word_wrap = True
            text_frame.clear()
            p = text_frame.paragraphs[0]
            p.text = text_content  # 所有内容放一段，让PPT自动根据宽度折行

            font = p.font
            font.size = font_size_pt
            font_family = styles.get('font-family', 'Arial')
            if font_family:
                font.name = font_family.strip().strip("'\"")
            color_tuple = self._parse_color(styles.get('color', 'black'))
            if color_tuple:
                font.color.rgb = RGBColor(*color_tuple)
            if styles.get('font-weight') == 'bold':
                font.bold = True
            text_align_style = styles.get('text-align')
            if text_align_style == 'center':
                p.alignment = PP_ALIGN.CENTER
            elif text_align_style == 'right':
                p.alignment = PP_ALIGN.RIGHT
            elif text_align_style == 'justify':
                p.alignment = PP_ALIGN.JUSTIFY
            else:
                p.alignment = PP_ALIGN.LEFT

    def walk_elements(self, element, slide, parent_left=0, parent_top=0):
        """
        递归遍历所有带 style 的元素，并累加父容器偏移，实现绝对定位。
        """
        styles = self._get_element_styles(element)
        # 只处理有 position: absolute 的元素
        if 'position' in styles and styles['position'] == 'absolute':
            left = parent_left + self._to_px(styles.get('left', '0px'))
            top = parent_top + self._to_px(styles.get('top', '0px'))
            self._add_element_to_slide_with_abs(slide, element, left, top)
        # 递归子元素
        for child in getattr(element, 'children', []):
            if isinstance(child, Tag):
                # 如果本元素有 position: absolute，则累加，否则父偏移不变
                next_left = parent_left
                next_top = parent_top
                if 'position' in styles and styles['position'] == 'absolute':
                    next_left = parent_left + self._to_px(styles.get('left', '0px'))
                    next_top = parent_top + self._to_px(styles.get('top', '0px'))
                self.walk_elements(child, slide, next_left, next_top)

    def convert(self, html_content: str):
        try:
            cleaned_html = re.sub(r'[\r\n\t]+', ' ', html_content)
            cleaned_html = re.sub(r'\s{2,}', ' ', cleaned_html)
            soup = BeautifulSoup(cleaned_html, 'html.parser')
            slide_containers = soup.find_all('div', class_='slide-page')
            if not slide_containers:
                slide = self.prs.slides.add_slide(self.blank_layout)
                self.walk_elements(soup.body, slide)
            for container in slide_containers:
                slide = self.prs.slides.add_slide(self.blank_layout)
                self.walk_elements(container, slide)
            self.prs.save(self.output_path)
            print(f"PPT文件已成功保存到 {self.output_path}")
        except Exception as e:
            print(f"转换过程中出错: {e}")
            raise
from PIL import ImageFont

def wrap_text_by_width(text, font_path, font_size, max_width_px):
    """
    按照像素宽度自动换行，返回自动断行后的文本列表（每行一个字符串）
    """
    font = ImageFont.truetype(font_path, int(font_size))
    lines = []
    for paragraph in text.split('\n'):
        buf = ''
        for char in paragraph:
            test_buf = buf + char
            width = font.getbbox(test_buf)[2] - font.getbbox(test_buf)[0]
            if width > max_width_px and buf:
                lines.append(buf)
                buf = char
            else:
                buf = test_buf
        if buf:
            lines.append(buf)
    return lines
```

使用时，初始化一个`LLMToPPTXConverter`实例，调用`convert`方法就可以。 

实测转化效果与原始html差别不大。

