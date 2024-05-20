# Personalized Image Generator for Academic Profiles

This project implements a personalized image generator for academic profiles, including details such as name, date of birth, grades, CGPA, and backlog history. It integrates VBA code for seamless sharing of generated images via WhatsApp, simplifying the process of sending academic profiles to individuals.

## Features

- **Customized Academic Profiles**: Generates personalized academic profiles for individuals based on provided data.
- **VBA Integration**: Utilizes VBA code to facilitate easy sharing of generated images via WhatsApp. The code automatically copies the image to the clipboard for quick pasting and sending.

## How It Works

1. **Data Input**:
   - Users provide input data, including name, date of birth, academic grades, CGPA, and backlog information.
   - ![](snapshots/Sample%20dataset.png)

2. **Sample Academic Profile Card**:
   - A sample academic profile card is generated based on the provided data.
   - ![](snapshots/Sample%20Grade%20card.png)

3. **Image Generation**:
   - The system uses VLOOKUP functions to create a grade card-like structure containing the individual's academic details.
   - ![](snapshots/Selection%20%of%20names%20.png)

4. **WhatsApp Integration**:
   - Clicking the generated image triggers the VBA code, which opens WhatsApp, redirects to the individual's number, and copies the image to the clipboard for sharing.
   - ![](snapshots/Sending%20in%20whatsapp.png)


## Usage

1. Input the required data for the individual's academic profile.
2. Select the person's name from the dropdown menu to generate the personalized image.
3. Click the image to automatically copy it to the clipboard.
4. Paste the image in WhatsApp and send it to the respective individual.

## Technologies Used

- Excel (especially Vlookup)
- VBA (for WhatsApp integration)


