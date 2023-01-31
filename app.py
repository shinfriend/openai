import os
import openai
from flask import Flask, redirect, render_template, request, url_for
from pptx import Presentation
from pptx.util import Inches, Pt
import urllib.request
import time

app = Flask(__name__)
openai.api_key = os.getenv("OPENAI_API_KEY")

@app.route("/", methods=("GET", "POST"))
def index():
    final_array=[]
    topics =""
    if request.method == "POST":
        topics = request.form["topics"]
        if topics:
           final_array.append(topics)
        else:
             final_array.append("")

        #Get the content
        response1 = openai.Completion.create(
        model="text-davinci-003",
        prompt=topics,  
        temperature=0.3,
        max_tokens=650,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
        )
        final_array.append(response1.choices[0].text)
        
        #Get the Image
        content = openai.Image.create(
        prompt=topics,
        n=1,
        size="512x512"
        )
        final_array.append(content['data'][0]['url'])

        #print("This is value", final_array[0])
        image_path = generate_powerpoint(final_array[0], final_array[1], final_array[2])
        print("This is image path:" + image_path)
        final_array.append(image_path)
        return redirect(url_for("index", result=final_array))  
    final_array1 = request.args.getlist("result")
   
    return render_template("index.html",final_array=final_array1)

def generate_prompt(topics):
    return """Suggest three names for an animal that is a superhero.
Animal: Cat
Names: Captain Sharpclaw, Agent Fluffball, The Incredible Feline
Animal: Dog
Names: Ruff the Protector, Wonder Canine, Sir Barks-a-Lot
Animal: {}
Names:""".format(
        topics.capitalize()
    )

def generate_powerpoint(title_txt,content,image_url):
    prs = Presentation()
    demo_ppt = ""
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = title_txt
    subtitle.text = content
    file_name = "image" + time.strftime("%Y%m%d-%H%M%S")
    image_path = download_image(image_url, 'static/images/', file_name)
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    left = top = Inches(1)
    pic = slide.shapes.add_picture(image_path, left, top)
    demo_ppt = "static/presentation/demo" +  time.strftime("%Y%m%d-%H%M%S") + ".pptx"
    prs.save(demo_ppt)
    return demo_ppt


def download_image(url, file_path, file_name):
    full_path = ""
    full_path = file_path + file_name + '.jpg'
    urllib.request.urlretrieve(url, full_path)
    return full_path