
# PowerPoint Gesture Control Presentation

## Overview

This project aims to create an interactive PowerPoint presentation controller using hand gestures captured through a webcam. By leveraging hand tracking technology, users can navigate through slides with simple gestures, enhancing the presentation experience without the need for traditional remote controls or clickers.

## Motivation

In today's fast-paced world, presentations are a critical part of sharing ideas and information. However, traditional methods of controlling slides can be cumbersome and limit the presenter’s ability to engage with the audience. Our motivation behind this project was to develop a more intuitive and seamless way to navigate presentations using natural hand gestures.

The increasing interest in virtual and augmented reality technologies has also highlighted the potential for gesture-based controls in various applications. By using computer vision and hand tracking, we can create a more interactive experience that not only makes presentations smoother but also showcases the possibilities of modern technology in education and business environments.

## Sample Video

[View Sample Videos on Google Drive](https://drive.google.com/file/d/1DQ9ijfiHRqt2zZ-WuV6Y0ADbh-HOl4uW/view)
<video width="600" controls autoplay loop>
  <source src="https://drive.google.com/file/d/1DQ9ijfiHRqt2zZ-WuV6Y0ADbh-HOl4uW/view" type="video/mp4">
  Your browser does not support the video tag.
</video>

## Features

- **Gesture Recognition**: Utilizes hand tracking to detect specific gestures for navigating slides:
  - Waving right with any hand advances to the next slide.
  - Waving left with any hand returns to the previous slide.
  - Making a fist (all fingers down) stops the slideshow.
  
- **Real-Time Interaction**: The application operates in real-time, allowing for immediate feedback as gestures are detected.

- **PowerPoint Integration**: Directly interacts with Microsoft PowerPoint, making it easy to control slides without needing third-party applications.

- **User-Friendly Interface**: A simple webcam setup makes the system accessible for users without technical expertise.

## Installation

To run this project, follow the steps below:

### Prerequisites

Make sure you have Python installed on your system. It's recommended to use a virtual environment to manage dependencies.

### Setup

1. **Clone the Repository**
   ```bash
   git clone <repository-url>
   cd <repository-directory>
   ```

2. **Create a Virtual Environment (Optional but recommended)**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```

3. **Install Required Packages**
   Create a file named `requirements.txt` in the project directory and include the following:
   ```plaintext
   opencv-python
   cvzone
   pywin32
   numpy
   ```
   Then install the dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. **Open PowerPoint Presentation**
   Ensure you have a PowerPoint presentation file ready. Update the script to point to your presentation file location.

   Also you need to change the location of the powerpoint in the main.py

5. **Run the Application**
   Execute the Python script:
   ```bash
   python your_script_name.py
   ```
   Replace `your_script_name.py` with the actual name of your Python file.

### Usage

1. **Position Yourself**: Make sure your webcam can capture your hand movements clearly. Position yourself at a reasonable distance from the camera.
   
2. **Control Presentation**: Use the specified hand gestures to navigate through the slides. 
   - Wave right to move to the next slide.
   - Wave left to go back to the previous slide.
   - Make a fist to stop the presentation.

3. **Exit**: To stop the application, either make a fist gesture to exit the slideshow or close the program window.

## Contributing

We welcome contributions to improve this project. If you have suggestions, ideas, or bug fixes, please fork the repository and submit a pull request.

## Future Enhancements

- **Additional Gestures**: Implement more gestures for controlling other features of the presentation, such as pausing or resuming.
- **Customization**: Allow users to customize gestures for different actions based on their preferences.
- **Multi-Hand Support**: Extend the application to support multiple hands for more complex gestures.
- **Cross-Platform Compatibility**: Explore making the application compatible with other presentation software besides PowerPoint.

## Conclusion

This project showcases how technology can be leveraged to create more interactive and engaging experiences in presentations. By utilizing hand gestures for navigation, we not only enhance the usability of presentations but also open the door to future innovations in the field of gesture recognition.
