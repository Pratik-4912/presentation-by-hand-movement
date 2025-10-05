import cv2
import mediapipe as mp
import pyautogui
import math
import time
import os
import sys
import os

# Path to your PPT file
ppt_path = "ppt1.pptx"

# Open PowerPoint presentation automatically
os.startfile(ppt_path)


# Optional Windows PowerPoint auto-start support
try:
    import pythoncom
    from win32com.client import Dispatch
    WIN32_AVAILABLE = True
except Exception:
    WIN32_AVAILABLE = False

# ---------- Configuration ----------
SWIPE_THRESHOLD = 0.20      
SWIPE_TIME = 0.5            
OPEN_HAND_FINGERS = 4       
COOLDOWN = 0.6              

# ---------- MediaPipe setup ----------
mp_hands = mp.solutions.hands
mp_draw = mp.solutions.drawing_utils
hands = mp_hands.Hands(max_num_hands=1, min_detection_confidence=0.6, min_tracking_confidence=0.6)

cap = cv2.VideoCapture(0)
if not cap.isOpened():
    print("Cannot open webcam. Exiting.")
    sys.exit(1)

last_action_time = 0.0
# For swipe detection: store (time, x) samples
samples = []

def fingers_extended(hand_landmarks):
    # Returns number of extended fingers (simple heuristic using y-coordinates for fingers)
    tips_ids = [4, 8, 12, 16, 20]  # thumb, index, middle, ring, pinky
    extended = 0
    lm = hand_landmarks.landmark
    # For fingers except thumb, compare tip y to pip y (lower y value means higher on screen when flipped)
    for tip_id in tips_ids[1:]:
        tip = lm[tip_id]
        pip = lm[tip_id - 2]
        if tip.y < pip.y:  # tip above pip = extended (note: image is flipped later)
            extended += 1
    # Thumb: compare tip x to ip x depending on hand orientation (rough heuristic)
    thumb_tip = lm[4]
    thumb_ip = lm[3]
    if abs(thumb_tip.x - thumb_ip.x) > 0.03:
        extended += 1
    return extended

def try_start_presentation(ppt_path):
    """Try to open the provided pptx file and start slideshow (Windows only)."""
    if not WIN32_AVAILABLE:
        print("win32com not available — automatic slideshow start disabled. Please open the PPTX and press F5 to start slideshow.")
        return False
    try:
        pythoncom.CoInitialize()
        ppt_app = Dispatch("PowerPoint.Application")
        ppt_app.Visible = True
        presentation = ppt_app.Presentations.Open(os.path.abspath(ppt_path))
        presentation.SlideShowSettings.Run()
        print("Presentation started via PowerPoint COM.")
        return True
    except Exception as e:
        print("Could not start slideshow via COM:", e)
        return False

def perform_action(action):
    global last_action_time
    now = time.time()
    if now - last_action_time < COOLDOWN:
        return
    last_action_time = now
    print("ACTION:", action)
    if action == "next":
        pyautogui.press('right')
    elif action == "prev":
        pyautogui.press('left')
    elif action == "pause":
        # toggle black screen / pause using 'b' (works in PowerPoint) or space to pause videos
        pyautogui.press('b')
    elif action == "resume":
        pyautogui.press('b')
    elif action == "start":
        # start slideshow by F5 if user didn't auto-start via COM
        pyautogui.press('f5')

print("Gesture Presentation Tool\nPress 'o' to open a PPTX file and (optionally) auto-start slideshow (Windows COM)." )
print("Press 'q' in the video window to quit.")

ppt_path = None

while True:
    ret, frame = cap.read()
    if not ret:
        break
    frame = cv2.flip(frame, 1)
    h, w, _ = frame.shape
    rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    result = hands.process(rgb)

    # Draw instructions overlay
    cv2.rectangle(frame, (0,0), (w,40), (0,0,0), -1)
    cv2.putText(frame, "o: Open PPTX  |  q: Quit", (10,25), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255,255,255), 2)

    if result.multi_hand_landmarks:
        hand_landmarks = result.multi_hand_landmarks[0]
        mp_draw.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)

        # compute normalized hand center x
        lm = hand_landmarks.landmark
        cx = sum([p.x for p in lm]) / len(lm)  # normalized 0..1
        cy = sum([p.y for p in lm]) / len(lm)

        # save sample for swipe detection
        samples.append((time.time(), cx))
        # remove old samples older than SWIPE_TIME
        now = time.time()
        samples[:] = [s for s in samples if now - s[0] <= SWIPE_TIME]

        # detect swipe: compare oldest x to newest x
        if len(samples) >= 2:
            dx = samples[-1][1] - samples[0][1]
            if abs(dx) > SWIPE_THRESHOLD:
                if dx > 0:
                    # swipe right (hand moved right) => next slide
                    perform_action("next")
                else:
                    perform_action("prev")
                samples.clear()

        # detect open hand for pause/resume: count extended fingers
        ext = fingers_extended(hand_landmarks)
        # show count
        cv2.putText(frame, f"Fingers: {ext}", (10, h-20), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0,255,0), 2)
        if ext >= OPEN_HAND_FINGERS:
            # open hand -> toggle pause (black screen)
            perform_action("pause")
            # small sleep to avoid immediate repeated toggles
            time.sleep(0.4)

    # show frame
    cv2.imshow("Gesture Presentation Tool", frame)
    key = cv2.waitKey(1) & 0xFF
    if key == ord('q'):
        break
    if key == ord('o'):
        # ask user to input path via console prompt
        print("Enter path to your PPTX file (or press Enter to cancel):")
        user_input = input().strip()
        if user_input:
            ppt_path = user_input
            print("Selected:", ppt_path)
            started = try_start_presentation(ppt_path)
            if not started:
                print("If COM auto-start failed, please open the PPTX manually and press F5 to start slideshow. This tool will control the slideshow once it is active.")

cap.release()
cv2.destroyAllWindows()
