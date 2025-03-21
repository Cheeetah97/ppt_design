from io import BytesIO
from copy import deepcopy

def find_shapes(prs, slide_index):
    slide = prs.slides[slide_index]
    for shape in slide.shapes:
        type = str(shape.shape_type).split(' ')[0]
        if type == 'TEXT_BOX':
            print(shape, f"Shape ID: {shape.shape_id}, Text: {shape.text}")
        elif type == 'PICTURE':
            print(shape, f"Shape ID: {shape.shape_id}, Picture: {shape.image}")
        elif type == 'TABLE':
            print(shape, f"Shape ID: {shape.shape_id}, Table: {shape.table}")
        else:
            print(shape, f"Shape ID: {shape.shape_id}")

def image_file_to_bytes(image_path):
    """
    Convert a local image file to bytes.
    
    Args:
        image_path (str): Path to the image file
        
    Returns:
        bytes: The image as bytes
    """
    with open(image_path, 'rb') as image_file:
        return image_file.read()

def duplicate_slide(presentation, source_slide):
    # Identify the source slide's layout
    slide_layout = source_slide.slide_layout

    # Create a new blank slide with the same layout
    new_slide = presentation.slides.add_slide(slide_layout)

    # Copy background (including images)
    source_bg = source_slide.background
    new_bg = new_slide.background
    
    if source_bg.fill.type == 1: # Solid fill
        try:
            new_bg.fill.solid()
            new_bg.fill.fore_color.rgb = source_bg.fill.fore_color.rgb
        except:
            pass
    
    elif source_bg.fill.type == 2:  # Picture fill
        # Extract image from source background
        img_blob = source_bg.fill.image.blob
        # Add image to new slide background
        new_bg.fill.picture()
        new_bg.fill.picture.image = presentation.slides[0].shapes.add_picture(img_blob, 0, 0, 0, 0).image  # Dummy shape to register image

    else:
        pass
    
    # Remove default shapes from the new slide (from the layout)
    for shape in new_slide.shapes:
        shape.element.getparent().remove(shape.element)

    # Deep copy all shapes from the source slide to the new slide
    for shape in source_slide.shapes:
        type = str(shape.shape_type).split(' ')[0]
        if type == 'PICTURE':
            left = shape.left
            top = shape.top
            width = shape.width
            height = shape.height
            image = shape.image
            image_stream = BytesIO(image.blob)
            new_slide.shapes.add_picture(image_stream, left, top, width=width, height=height)
        else:
            new_shape = deepcopy(shape.element)
            new_slide.shapes._spTree.insert_element_before(new_shape, 'p:extLst')

def add_title_slide(prs, **kwargs):

    # Duplicate the title slide to create a new slide
    original_slide = prs.slides[kwargs['template']['id'] - 1]
    duplicate_slide(prs, original_slide)
    new_slide = prs.slides[-1]  # Get the newly created slide

    # Process each element defined in the template
    for element in kwargs['template']['elements']:
        element_id = element['id']
        element_type = element['type']

        for shape in new_slide.shapes:
            if shape.shape_id == element_id:

                if element_type == 'title':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['title']

                elif element_type == 'credit_line':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['credit_line']

                elif element_type == 'date':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['date']

                elif element_type == 'image':
                    if kwargs['image']:
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height

                        shape._element.getparent().remove(shape._element)

                        if not isinstance(kwargs['image'], BytesIO):
                            image_stream = BytesIO(kwargs['image'])
                        else:
                            image_stream = kwargs['image']

                        new_slide.shapes.add_picture(image_stream, left, top, width, height)

                break

def add_paragraphs_slide(prs, **kwargs):

    # Duplicate the title slide to create a new slide
    original_slide = prs.slides[kwargs['template']['id'] - 1]
    duplicate_slide(prs, original_slide)
    new_slide = prs.slides[-1]  # Get the newly created slide

    # Process each element defined in the template
    for element in kwargs['template']['elements']:
        element_id = element['id']
        element_type = element['type']

        for shape in new_slide.shapes:
            if shape.shape_id == element_id:

                if element_type == 'title':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['title']
                        
                elif element_type == 'paragraphs':
                    frame = shape.text_frame

                    while len(frame.paragraphs) > len(kwargs['paragraphs']):
                        p = frame.paragraphs[-1]
                        frame._element.remove(p._element)
                    
                    for i, paragraph_text in enumerate(kwargs['paragraphs']):
                        if i < len(frame.paragraphs):
                            frame.paragraphs[i].runs[0].text = paragraph_text
                            # if i != len(kwargs['paragraphs']) - 1:
                            #     frame.paragraphs[i].runs[0].font.name = frame.paragraphs[0].runs[0].font.name
                            #     frame.paragraphs[i].runs[0].font.size = frame.paragraphs[0].runs[0].font.size
                            #     frame.paragraphs[i].add_line_break()
                        else:
                            # Add a line break before adding a new paragraph
                            if i > 0:  # Add line break only if it's not the first paragraph
                                frame.paragraphs[-1].add_line_break()
                            # Add new paragraph with the same formatting
                            p = frame.add_paragraph()
                            p.text = paragraph_text
                            if frame.paragraphs:  # Copy formatting from the first paragraph
                                p.runs[0].font.name = frame.paragraphs[0].runs[0].font.name
                                p.runs[0].font.size = frame.paragraphs[0].runs[0].font.size
                                p.runs[0].font.color.rgb = frame.paragraphs[0].runs[0].font.color.rgb

                break

def add_paragraphs_slide_with_icon(prs, **kwargs):

    # Duplicate the title slide to create a new slide
    original_slide = prs.slides[kwargs['template']['id'] - 1]
    duplicate_slide(prs, original_slide)
    new_slide = prs.slides[-1]  # Get the newly created slide

    # Process each element defined in the template
    for element in kwargs['template']['elements']:
        element_id = element['id']
        element_type = element['type']

        for shape in new_slide.shapes:
            if shape.shape_id == element_id:

                if element_type == 'title':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['title']
                        
                elif element_type == 'paragraphs':
                    frame = shape.text_frame

                    while len(frame.paragraphs) > len(kwargs['paragraphs']):
                        p = frame.paragraphs[-1]
                        frame._element.remove(p._element)
                    
                    for i, paragraph_text in enumerate(kwargs['paragraphs']):
                        if i < len(frame.paragraphs):
                            frame.paragraphs[i].runs[0].text = paragraph_text
                        else:
                            # Add a line break before adding a new paragraph
                            if i > 0:  # Add line break only if it's not the first paragraph
                                frame.paragraphs[-1].add_line_break()
                            # Add new paragraph with the same formatting
                            p = frame.add_paragraph()
                            p.text = paragraph_text
                            if frame.paragraphs:  # Copy formatting from the first paragraph
                                p.runs[0].font.name = frame.paragraphs[0].runs[0].font.name
                                p.runs[0].font.size = frame.paragraphs[0].runs[0].font.size
                                p.runs[0].font.color.rgb = frame.paragraphs[0].runs[0].font.color.rgb

                elif element_type == 'icon':
                    if kwargs['icon']:
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height

                        shape._element.getparent().remove(shape._element)

                        if not isinstance(kwargs['icon'], BytesIO):
                            image_stream = BytesIO(kwargs['icon'])
                        else:
                            image_stream = kwargs['icon']

                        new_slide.shapes.add_picture(image_stream, left, top, width, height)

                break

def add_paragraphs_slide_with_image(prs, **kwargs):

    # Duplicate the title slide to create a new slide
    original_slide = prs.slides[kwargs['template']['id'] - 1]
    duplicate_slide(prs, original_slide)
    new_slide = prs.slides[-1]  # Get the newly created slide

    # Process each element defined in the template
    for element in kwargs['template']['elements']:
        element_id = element['id']
        element_type = element['type']

        for shape in new_slide.shapes:
            if shape.shape_id == element_id:

                if element_type == 'title':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['title']
                        
                elif element_type == 'paragraphs':
                    frame = shape.text_frame

                    while len(frame.paragraphs) > len(kwargs['paragraphs']):
                        p = frame.paragraphs[-1]
                        frame._element.remove(p._element)
                    
                    for i, paragraph_text in enumerate(kwargs['paragraphs']):
                        if i < len(frame.paragraphs):
                            frame.paragraphs[i].runs[0].text = paragraph_text
                        else:
                            # Add a line break before adding a new paragraph
                            if i > 0:  # Add line break only if it's not the first paragraph
                                frame.paragraphs[-1].add_line_break()
                            # Add new paragraph with the same formatting
                            p = frame.add_paragraph()
                            p.text = paragraph_text
                            if frame.paragraphs:  # Copy formatting from the first paragraph
                                p.runs[0].font.name = frame.paragraphs[0].runs[0].font.name
                                p.runs[0].font.size = frame.paragraphs[0].runs[0].font.size
                                p.runs[0].font.color.rgb = frame.paragraphs[0].runs[0].font.color.rgb

                elif element_type == 'image':
                    if kwargs['image']:
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height

                        shape._element.getparent().remove(shape._element)

                        if not isinstance(kwargs['image'], BytesIO):
                            image_stream = BytesIO(kwargs['image'])
                        else:
                            image_stream = kwargs['image']

                        new_slide.shapes.add_picture(image_stream, left, top, width, height)

                break

def add_unordered_bullets_slide(prs, **kwargs):

    # Duplicate the title slide to create a new slide
    original_slide = prs.slides[kwargs['template']['id'] - 1]
    duplicate_slide(prs, original_slide)
    new_slide = prs.slides[-1]  # Get the newly created slide

    # Process each element defined in the template
    for element in kwargs['template']['elements']:
        element_id = element['id']
        element_type = element['type']

        for shape in new_slide.shapes:
            if shape.shape_id == element_id:

                if element_type == 'title':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['title']
                        
                elif element_type == 'bullets':
                    frame = shape.text_frame

                    while len(frame.paragraphs) > len(kwargs['bullets']):
                        p = frame.paragraphs[-1]
                        frame._element.remove(p._element)
                    
                    for i, paragraph_text in enumerate(kwargs['bullets']):
                        if i < len(frame.paragraphs):
                            frame.paragraphs[i].runs[0].text = paragraph_text
                        else:
                            # Add a line break before adding a new paragraph
                            if i > 0:  # Add line break only if it's not the first paragraph
                                frame.paragraphs[-1].add_line_break()

                            new_p_element = deepcopy(frame.paragraphs[-1]._element)
                            frame._element.insert(frame._element.index(frame.paragraphs[-1]._element) + 1, new_p_element)
                            p = frame.paragraphs[-1]
                            p.text = paragraph_text
                            p.level = 0

                            if frame.paragraphs:  # Copy formatting from the first paragraph
                                p.runs[0].font.name = frame.paragraphs[0].runs[0].font.name
                                p.runs[0].font.size = frame.paragraphs[0].runs[0].font.size
                                p.runs[0].font.color.rgb = frame.paragraphs[0].runs[0].font.color.rgb

                break

def add_unordered_bullets_slide_with_icon(prs, **kwargs):

    # Duplicate the title slide to create a new slide
    original_slide = prs.slides[kwargs['template']['id'] - 1]
    duplicate_slide(prs, original_slide)
    new_slide = prs.slides[-1]  # Get the newly created slide

    # Process each element defined in the template
    for element in kwargs['template']['elements']:
        element_id = element['id']
        element_type = element['type']

        for shape in new_slide.shapes:
            if shape.shape_id == element_id:

                if element_type == 'title':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['title']
                        
                elif element_type == 'bullets':
                    frame = shape.text_frame

                    while len(frame.paragraphs) > len(kwargs['bullets']):
                        p = frame.paragraphs[-1]
                        frame._element.remove(p._element)
                    
                    for i, paragraph_text in enumerate(kwargs['bullets']):
                        if i < len(frame.paragraphs):
                            frame.paragraphs[i].runs[0].text = paragraph_text
                        else:
                            # Add a line break before adding a new paragraph
                            if i > 0:  # Add line break only if it's not the first paragraph
                                frame.paragraphs[-1].add_line_break()

                            new_p_element = deepcopy(frame.paragraphs[-1]._element)
                            frame._element.insert(frame._element.index(frame.paragraphs[-1]._element) + 1, new_p_element)
                            p = frame.paragraphs[-1]
                            p.text = paragraph_text
                            p.level = 0

                            if frame.paragraphs:  # Copy formatting from the first paragraph
                                p.runs[0].font.name = frame.paragraphs[0].runs[0].font.name
                                p.runs[0].font.size = frame.paragraphs[0].runs[0].font.size
                                p.runs[0].font.color.rgb = frame.paragraphs[0].runs[0].font.color.rgb

                elif element_type == 'icon':
                    if kwargs['icon']:
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height

                        shape._element.getparent().remove(shape._element)

                        if not isinstance(kwargs['icon'], BytesIO):
                            image_stream = BytesIO(kwargs['icon'])
                        else:
                            image_stream = kwargs['icon']

                        new_slide.shapes.add_picture(image_stream, left, top, width, height)

                break

def add_unordered_bullets_slide_with_image(prs, **kwargs):

    # Duplicate the title slide to create a new slide
    original_slide = prs.slides[kwargs['template']['id'] - 1]
    duplicate_slide(prs, original_slide)
    new_slide = prs.slides[-1]  # Get the newly created slide

    # Process each element defined in the template
    for element in kwargs['template']['elements']:
        element_id = element['id']
        element_type = element['type']

        for shape in new_slide.shapes:
            if shape.shape_id == element_id:

                if element_type == 'title':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['title']
                        
                elif element_type == 'bullets':
                    frame = shape.text_frame

                    while len(frame.paragraphs) > len(kwargs['bullets']):
                        p = frame.paragraphs[-1]
                        frame._element.remove(p._element)
                    
                    for i, paragraph_text in enumerate(kwargs['bullets']):
                        if i < len(frame.paragraphs):
                            frame.paragraphs[i].runs[0].text = paragraph_text
                        else:
                            # Add a line break before adding a new paragraph
                            if i > 0:  # Add line break only if it's not the first paragraph
                                frame.paragraphs[-1].add_line_break()

                            new_p_element = deepcopy(frame.paragraphs[-1]._element)
                            frame._element.insert(frame._element.index(frame.paragraphs[-1]._element) + 1, new_p_element)
                            p = frame.paragraphs[-1]
                            p.text = paragraph_text
                            p.level = 0

                            if frame.paragraphs:  # Copy formatting from the first paragraph
                                p.runs[0].font.name = frame.paragraphs[0].runs[0].font.name
                                p.runs[0].font.size = frame.paragraphs[0].runs[0].font.size
                                p.runs[0].font.color.rgb = frame.paragraphs[0].runs[0].font.color.rgb

                elif element_type == 'image':
                    if kwargs['image']:
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height

                        shape._element.getparent().remove(shape._element)

                        if not isinstance(kwargs['image'], BytesIO):
                            image_stream = BytesIO(kwargs['image'])
                        else:
                            image_stream = kwargs['image']

                        new_slide.shapes.add_picture(image_stream, left, top, width, height)

                break

def add_ordered_bullets_slide(prs, **kwargs):

    # Duplicate the title slide to create a new slide
    original_slide = prs.slides[kwargs['template']['id'] - 1]
    duplicate_slide(prs, original_slide)
    new_slide = prs.slides[-1]  # Get the newly created slide

    # Process each element defined in the template
    for element in kwargs['template']['elements']:
        element_id = element['id']
        element_type = element['type']

        for shape in new_slide.shapes:
            if shape.shape_id == element_id:

                if element_type == 'title':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['title']
                        
                elif element_type == 'bullets':
                    frame = shape.text_frame

                    while len(frame.paragraphs) > len(kwargs['bullets']):
                        p = frame.paragraphs[-1]
                        frame._element.remove(p._element)
                    
                    for i, paragraph_text in enumerate(kwargs['bullets']):
                        if i < len(frame.paragraphs):
                            frame.paragraphs[i].runs[0].text = paragraph_text
                        else:
                            # Add a line break before adding a new paragraph
                            if i > 0:  # Add line break only if it's not the first paragraph
                                frame.paragraphs[-1].add_line_break()

                            new_p_element = deepcopy(frame.paragraphs[-1]._element)
                            frame._element.insert(frame._element.index(frame.paragraphs[-1]._element) + 1, new_p_element)
                            p = frame.paragraphs[-1]
                            p.text = paragraph_text
                            p.level = 0

                            if frame.paragraphs:  # Copy formatting from the first paragraph
                                p.runs[0].font.name = frame.paragraphs[0].runs[0].font.name
                                p.runs[0].font.size = frame.paragraphs[0].runs[0].font.size
                                p.runs[0].font.color.rgb = frame.paragraphs[0].runs[0].font.color.rgb

                break

def add_ordered_bullets_slide_with_icon(prs, **kwargs):

    # Duplicate the title slide to create a new slide
    original_slide = prs.slides[kwargs['template']['id'] - 1]
    duplicate_slide(prs, original_slide)
    new_slide = prs.slides[-1]  # Get the newly created slide

    # Process each element defined in the template
    for element in kwargs['template']['elements']:
        element_id = element['id']
        element_type = element['type']

        for shape in new_slide.shapes:
            if shape.shape_id == element_id:

                if element_type == 'title':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['title']
                        
                elif element_type == 'bullets':
                    frame = shape.text_frame

                    while len(frame.paragraphs) > len(kwargs['bullets']):
                        p = frame.paragraphs[-1]
                        frame._element.remove(p._element)
                    
                    for i, paragraph_text in enumerate(kwargs['bullets']):
                        if i < len(frame.paragraphs):
                            frame.paragraphs[i].runs[0].text = paragraph_text
                        else:
                            # Add a line break before adding a new paragraph
                            if i > 0:  # Add line break only if it's not the first paragraph
                                frame.paragraphs[-1].add_line_break()

                            new_p_element = deepcopy(frame.paragraphs[-1]._element)
                            frame._element.insert(frame._element.index(frame.paragraphs[-1]._element) + 1, new_p_element)
                            p = frame.paragraphs[-1]
                            p.text = paragraph_text
                            p.level = 0

                            if frame.paragraphs:  # Copy formatting from the first paragraph
                                p.runs[0].font.name = frame.paragraphs[0].runs[0].font.name
                                p.runs[0].font.size = frame.paragraphs[0].runs[0].font.size
                                p.runs[0].font.color.rgb = frame.paragraphs[0].runs[0].font.color.rgb

                elif element_type == 'icon':
                    if kwargs['icon']:
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height

                        shape._element.getparent().remove(shape._element)

                        if not isinstance(kwargs['icon'], BytesIO):
                            image_stream = BytesIO(kwargs['icon'])
                        else:
                            image_stream = kwargs['icon']

                        new_slide.shapes.add_picture(image_stream, left, top, width, height)

                break

def add_ordered_bullets_slide_with_image(prs, **kwargs):

    # Duplicate the title slide to create a new slide
    original_slide = prs.slides[kwargs['template']['id'] - 1]
    duplicate_slide(prs, original_slide)
    new_slide = prs.slides[-1]  # Get the newly created slide

    # Process each element defined in the template
    for element in kwargs['template']['elements']:
        element_id = element['id']
        element_type = element['type']

        for shape in new_slide.shapes:
            if shape.shape_id == element_id:

                if element_type == 'title':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['title']
                        
                elif element_type == 'bullets':
                    frame = shape.text_frame

                    while len(frame.paragraphs) > len(kwargs['bullets']):
                        p = frame.paragraphs[-1]
                        frame._element.remove(p._element)
                    
                    for i, paragraph_text in enumerate(kwargs['bullets']):
                        if i < len(frame.paragraphs):
                            frame.paragraphs[i].runs[0].text = paragraph_text
                        else:
                            # Add a line break before adding a new paragraph
                            if i > 0:  # Add line break only if it's not the first paragraph
                                frame.paragraphs[-1].add_line_break()

                            new_p_element = deepcopy(frame.paragraphs[-1]._element)
                            frame._element.insert(frame._element.index(frame.paragraphs[-1]._element) + 1, new_p_element)
                            p = frame.paragraphs[-1]
                            p.text = paragraph_text
                            p.level = 0

                            if frame.paragraphs:  # Copy formatting from the first paragraph
                                p.runs[0].font.name = frame.paragraphs[0].runs[0].font.name
                                p.runs[0].font.size = frame.paragraphs[0].runs[0].font.size
                                p.runs[0].font.color.rgb = frame.paragraphs[0].runs[0].font.color.rgb

                elif element_type == 'image':
                    if kwargs['image']:
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height

                        shape._element.getparent().remove(shape._element)

                        if not isinstance(kwargs['image'], BytesIO):
                            image_stream = BytesIO(kwargs['image'])
                        else:
                            image_stream = kwargs['image']

                        new_slide.shapes.add_picture(image_stream, left, top, width, height)

                break

def add_three_point_feature_slide(prs, **kwargs):

    # Duplicate the title slide to create a new slide
    original_slide = prs.slides[kwargs['template']['id'] - 1]
    duplicate_slide(prs, original_slide)
    new_slide = prs.slides[-1]  # Get the newly created slide

    # Process each element defined in the template
    for element in kwargs['template']['elements']:
        element_id = element['id']
        element_type = element['type']

        for shape in new_slide.shapes:
            if shape.shape_id == element_id:

                if element_type == 'title':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['title']
                
                elif element_type == 'heading_1':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['heading_1']
                
                elif element_type == 'heading_2':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['heading_2']
                
                elif element_type == 'heading_3':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['heading_3']
                
                elif element_type == 'content_1':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['content_1']
                
                elif element_type == 'content_2':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['content_2']
                
                elif element_type == 'content_3':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['content_3']
                
                break

def add_three_point_feature_slide_with_icons(prs, **kwargs):

    # Duplicate the title slide to create a new slide
    original_slide = prs.slides[kwargs['template']['id'] - 1]
    duplicate_slide(prs, original_slide)
    new_slide = prs.slides[-1]  # Get the newly created slide

    # Process each element defined in the template
    for element in kwargs['template']['elements']:
        element_id = element['id']
        element_type = element['type']

        for shape in new_slide.shapes:
            if shape.shape_id == element_id:

                if element_type == 'title':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['title']
                
                elif element_type == 'heading_1':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['heading_1']
                
                elif element_type == 'heading_2':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['heading_2']
                
                elif element_type == 'heading_3':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['heading_3']
                
                elif element_type == 'content_1':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['content_1']
                
                elif element_type == 'content_2':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['content_2']
                
                elif element_type == 'content_3':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['content_3']
                
                elif element_type == 'icon_1':
                    if kwargs['icon_1']:
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height

                        shape._element.getparent().remove(shape._element)

                        if not isinstance(kwargs['icon_1'], BytesIO):
                            image_stream = BytesIO(kwargs['icon_1'])
                        else:
                            image_stream = kwargs['icon_1']

                        new_slide.shapes.add_picture(image_stream, left, top, width, height)
                
                elif element_type == 'icon_2':
                    if kwargs['icon_2']:
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height

                        shape._element.getparent().remove(shape._element)

                        if not isinstance(kwargs['icon_2'], BytesIO):
                            image_stream = BytesIO(kwargs['icon_2'])
                        else:
                            image_stream = kwargs['icon_2']

                        new_slide.shapes.add_picture(image_stream, left, top, width, height)
                
                elif element_type == 'icon_3':
                    if kwargs['icon_3']:
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height

                        shape._element.getparent().remove(shape._element)

                        if not isinstance(kwargs['icon_3'], BytesIO):
                            image_stream = BytesIO(kwargs['icon_3'])
                        else:
                            image_stream = kwargs['icon_3']

                        new_slide.shapes.add_picture(image_stream, left, top, width, height)
                
                break

def add_table_slide(prs, **kwargs):

    # Duplicate the title slide to create a new slide
    original_slide = prs.slides[kwargs['template']['id'] - 1]
    duplicate_slide(prs, original_slide)
    new_slide = prs.slides[-1]  # Get the newly created slide

    # Process each element defined in the template
    for element in kwargs['template']['elements']:
        element_id = element['id']
        element_type = element['type']

        for shape in new_slide.shapes:
            if shape.shape_id == element_id:

                if element_type == 'title':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['title']
                
                elif element_type == 'table':
                    table = shape.table

                    # Get original table dimensions and total size BEFORE any modifications
                    original_total_width = sum(grid_col.w for grid_col in table._tbl.tblGrid.gridCol_lst)
                    original_total_height = sum(tr.h for tr in table._tbl.tr_lst)

                    if len(kwargs['rows']) > 5:
                        kwargs['rows'] = kwargs['rows'][:5]
                    
                    if len(kwargs['colnames']) > 5:
                        kwargs['colnames'] = kwargs['colnames'][:5]

                    # Calculate required dimensions
                    required_rows = len(kwargs['rows']) + 1  # +1 for header row
                    required_cols = len(kwargs['colnames'])

                    # Delete excess rows (working from bottom up)
                    while len(table.rows) > required_rows:
                        tr = table._tbl.tr_lst[-1]  # Get last row element
                        table._tbl.remove(tr)       # Remove last row from XML

                    # Delete excess columns (working from right to left)
                    while len(table.columns) > required_cols:
                        # Remove grid column (defines column structure)
                        grid_col = table._tbl.tblGrid.gridCol_lst[-1]
                        table._tbl.tblGrid.remove(grid_col)
                        
                        # Remove cells from each row
                        for row in table.rows:
                            tc = row._tr.tc_lst[-1]  # Get last cell in row
                            row._tr.remove(tc)
                    
                    # Evenly distribute column widths
                    if required_cols > 0:
                        new_col_width = int(original_total_width / required_cols)
                        for grid_col in table._tbl.tblGrid.gridCol_lst:
                            grid_col.w = new_col_width

                    # Evenly distribute row heights
                    if required_rows > 0:
                        new_row_height = int(original_total_height / required_rows)
                        for tr in table._tbl.tr_lst:
                            tr.h = new_row_height

                    # Write column headers
                    for col_idx, header in enumerate(kwargs['colnames']):
                        cell = table.cell(0, col_idx)
                        for p in cell.text_frame.paragraphs:
                            p.runs[0].text = str(header)

                    # Write data rows (starting from row 1)
                    for row_idx, row_data in enumerate(kwargs['rows'], start=1):
                        for col_idx, cell_value in enumerate(row_data):
                            cell = table.cell(row_idx, col_idx)
                            for p in cell.text_frame.paragraphs:
                                p.runs[0].text = str(cell_value)
                break


def add_thankyou_slide(prs, **kwargs):

    # Duplicate the title slide to create a new slide
    original_slide = prs.slides[kwargs['template']['id'] - 1]
    duplicate_slide(prs, original_slide)
    new_slide = prs.slides[-1]  # Get the newly created slide

    # Process each element defined in the template
    for element in kwargs['template']['elements']:
        element_id = element['id']
        element_type = element['type']

        for shape in new_slide.shapes:
            if shape.shape_id == element_id:

                if element_type == 'closing_statement':
                    shape.text_frame.paragraphs[0].runs[0].text = kwargs['closing_statement']
                
                break
                

def delete_slides_by_index(prs, slide_indexes):
    
    # Sort slide indexes in descending order to avoid index shifting issues
    slide_indexes_sorted = sorted(slide_indexes, reverse = True)

    # Delete slides by index
    for index in slide_indexes_sorted:
        if 0 <= index < len(prs.slides):
            xml_slides = prs.slides._sldIdLst  # Get the list of slide IDs
            xml_slides.remove(xml_slides[index])  # Remove the slide by index