import keyboard

def main_loop():
    while True:
        # Your loop logic here
        print("Working...")

        # Check for 'q' key press to break the loop
        if keyboard.is_pressed('q'):
            print("Loop terminated by user.")
            break

if __name__ == "__main__":
    main_loop()

    