import aspose.slides as slides
import aspose.slides.animation as anim
import pandas as pd

TRIGGER_TYPE_MAP = {
    anim.EffectTriggerType.AFTER_PREVIOUS: "After Previous",
    anim.EffectTriggerType.WITH_PREVIOUS: "With Previous",
    anim.EffectTriggerType.ON_CLICK: "On Click"
}

def get_animation_type(effect, shape):
    # If shape is a video and effect is None, label it as Play
    if isinstance(shape, slides.VideoFrame):
        return "Play"

    if effect.type == anim.EffectType.FADE:
        return "Fade"
    elif effect.type == anim.EffectType.WIPE:
        subtype = effect.subtype
        if subtype == anim.EffectSubtype.LEFT:
            return "Wipe Left to Right"
        elif subtype == anim.EffectSubtype.RIGHT:
            return "Wipe Right to Left"
        elif subtype == anim.EffectSubtype.TOP:
            return "Wipe Bottom to Top"
        elif subtype == anim.EffectSubtype.BOTTOM:
            return "Wipe Top to Bottom"
        else:
            return "Wipe"
    return "Unknown"

def run_animation_qc(pptx_path):
    pres = slides.Presentation(pptx_path)
    data = []

    for slide in pres.slides:
        for shape in slide.shapes:
            shape_name = shape.name if shape.name else "Unnamed Shape"
            text = ""
            if hasattr(shape, "text_frame") and shape.text_frame and shape.text_frame.text:
                text = shape.text_frame.text.strip()
            else:
                text = "No Text"

            effects = slide.timeline.main_sequence.get_effects_by_shape(shape)

            if effects:
                for effect in effects:
                    anim_type = get_animation_type(effect, shape)
                    delay = round(effect.timing.trigger_delay_time, 2)
                    trigger_type = TRIGGER_TYPE_MAP.get(effect.timing.trigger_type, "Unknown")
                    data.append({
                        "Slide": slide.slide_number,
                        "Shape Name / Table Cell": shape_name,
                        "Text": text,
                        "Animation Type": anim_type,
                        "Delay (sec)": delay,
                        "Trigger Type": trigger_type
                    })
            else:
                # If no effect, check if it's a video to label as Play
                anim_type = "Play" if isinstance(shape, slides.VideoFrame) else "None"
                data.append({
                    "Slide": slide.slide_number,
                    "Shape Name / Table Cell": shape_name,
                    "Text": text,
                    "Animation Type": anim_type,
                    "Delay (sec)": "",
                    "Trigger Type": ""
                })

    return pd.DataFrame(data)
