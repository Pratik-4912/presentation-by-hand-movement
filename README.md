Gesture Presentation Tool 🎬🖐️

Control your PowerPoint presentations using hand gestures — no keyboard or clicker needed!

Features

Next Slide: Swipe your hand right

Previous Slide: Swipe your hand left

Pause/Resume Presentation: Show an open palm

Optional Auto-Start: Automatically start PowerPoint slideshow on Windows

Real-Time Hand Tracking: Powered by MediaPipe and OpenCV

Tech Stack

Python 3.x

OpenCV: Webcam capture and video processing

MediaPipe: Hand tracking and gesture recognition

PyAutoGUI: Keyboard simulation for slide control

Windows COM Support (Optional): Auto-start PPT slideshow

Installation
git clone <your-repo-url>
cd gesture-presentation-tool
pip install -r requirements.txt

Usage

Run the script:

python app.py


When prompted, select your PPTX file or press Enter to use an already open presentation.

Control slides using hand gestures in front of your webcam.

How It Works

Captures video from your webcam

Detects your hand and tracks landmarks

Recognizes gestures (swipes, open hand)

Sends keyboard commands to PowerPoint to navigate slides






