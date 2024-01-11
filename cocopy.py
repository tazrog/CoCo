#Written by R Scharmen
#Designed to save programs via cassette interface and load them in the CoCo.

import wave
import pyaudio
from array import array
import os
import sys
import keyboard
import time

# Make sure directories are set up.
BASE_DIR = os.path.expanduser('~//Documents//CoCo')
#BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if not os.path.exists(BASE_DIR+"//Programs"):
    os.makedirs(BASE_DIR+"//Programs") 

def record():
    print ("Welcome to CoCo python tape player and recorder.")
    print ("For best results, please set PC sound volume between 54 to 58.")
    print (("The program's file location is at %s") % BASE_DIR+  "\programs")
    print ("")
    print ("RECORDING SECTION")
    
    FORMAT=pyaudio.paInt16
    CHANNELS=1
    RATE=9600
    CHUNK=1024
    RECORD_SECONDS=1
    wavitems = os.listdir(BASE_DIR+"\\programs")

    wavList = [wav for wav in wavitems if wav.endswith(".wav")]

    for fcnt, wavName in enumerate(wavList, 1):
        sys.stdout.write("[%d] %s\n\r" % (fcnt, wavName))
        print ("")
    print ("")
    if len(os.listdir(BASE_DIR+"\programs")) < 1: 
        print ("[9999] New File Name")
    else:
        print ("[9999] New File Name")
        print ("[99:] Main Menu")
    
    print ("Type in your option or a number of a file you will like to override. Do not include .wav extension.")
    tapename = int(input())
      
    if (tapename == 99):
        main()
    
    if (tapename == 9999):
        newname = input("Type new file name. Do not include .wav extension. >")
        rfn = newname +".wav"
    else:
        tapename = tapename -1
        rfn = (wavList[tapename])
        option = str(input ("Are you sure you wany to override file "+ rfn +" [y/n]? >"))   
        if (option != "y"):
            print ("File override canceled")
            time.sleep(1)
            record()   
    
    selection = input ("Enter the CSAVE command on the CoCo and press any key on the PC to start recording. \nAfter PC is recording execute CSAVE on PC. Enter x to exit. ")
    if (selection == "x"):
        print ("Program Exited")
        main()

    FILE_NAME=(BASE_DIR+"\\programs\\" + rfn)
    os.system('cls')
    print ("RECORDING TAPE...")
    print ("Press [q] to stop recording. There will be a short delay in stopping.") 
    print (("The program's file location is at %s") % BASE_DIR+  "\programs")
    print (rfn)
    audio=pyaudio.PyAudio() 
    stream=audio.open(format=FORMAT,channels=CHANNELS, 
        rate=RATE,
        input=True,
        frames_per_buffer=CHUNK)

    #start recording     
    frames=[] 
    
    while True:                 
        
        for i in range(0,int(RATE/CHUNK*RECORD_SECONDS)):
            if keyboard.is_pressed('q'):  # if key 'q' is pressed 
                print ("Stopping the recording")
                data=stream.read(CHUNK)
                frames.append(data)  
                stream.stop_stream()
                stream.close()
                audio.terminate()

                #end of recording
                stream.stop_stream()
                stream.close()
                audio.terminate()

                #writing to file
                wavfile=wave.open(FILE_NAME,'wb')
                wavfile.setnchannels(CHANNELS)
                wavfile.setsampwidth(audio.get_sample_size(FORMAT))
                wavfile.setframerate(RATE)
                wavfile.writeframes(b''.join(frames))#append frames recorded to file
                wavfile.close()
                os.system('cls')
                print ("Tape recording done for file named:") 
                print (FILE_NAME)
                time.sleep (1.5)
                os.system('cls')
                main()   
            data=stream.read(CHUNK)
            frames.append(data)    
    os.system('cls')    
    #end of recording
    stream.stop_stream()
    stream.close()
    audio.terminate()

    #writing to file
    wavfile=wave.open(FILE_NAME,'wb')
    wavfile.setnchannels(CHANNELS)
    wavfile.setsampwidth(audio.get_sample_size(FORMAT))
    wavfile.setframerate(RATE)
    wavfile.writeframes(b''.join(frames))#append frames recorded to file
    wavfile.close()
    print ("Tape recording done for file named:") 
    print (FILE_NAME)
    main()

def playtape():
    os.system('cls')  
    print ("Welcome to CoCo python tape player and recorder.")
    print ("For best results, please set PC sound volume between 54 to 58.")
    print (("The program's file location is at %s") % BASE_DIR+  "\programs")
    print ("")
    print ("PLAYING TAPE SECTION")
    #check for empty directory
    if len(os.listdir(BASE_DIR+"\programs")) < 1: 
        print ("")
        print("There are no files to play. Please record some.")
        time.sleep(2)
        main()
    chunk = 1024       
    items = os.listdir(BASE_DIR+"\programs")
    fileList = [name for name in items if name.endswith(".wav")]
    for fcnt, fileName in enumerate(fileList, 1):
        sys.stdout.write("[%d] %s\n\r" % (fcnt, fileName))
    print ("")
    print ("[99] Main Menu")
    
    choice = int(input("Select WAV file[1-%s]: " % fcnt))-1
    if (choice > 97):
        main()
    print ("Execute CLOAD on the CoCo and press enter when ready on the PC. Press 99 to exit to menu ")
    print ("")
    selection = input("> ")
    print ("PLAYING... ")
    
    print(fileList[choice])
    fn = (fileList[choice])
    wavplay = BASE_DIR + "\\programs\\"+ fn
    f = wave.open(wavplay,"rb")
    p = pyaudio.PyAudio()  
   
    
     
    stream = p.open(format = p.get_format_from_width(f.getsampwidth()),  
        channels = f.getnchannels(),  
        rate = f.getframerate(),  
        output = True,
        )
    
    print (f.getframerate())
    
  
    data = f.readframes(chunk)  
    
    #play stream  
    while data:  
        stream.write(data)  
        data = f.readframes(chunk)  
    
    #stop stream  
    stream.stop_stream()  
    stream.close()  

    #close PyAudio  
    p.terminate() 
    print ("Tape Complete")
    main()

def list():
    os.system('cls')
    print ("Welcome to CoCo python tape player and recorder.")
    print ("For best results, please set PC sound volume between 54 to 58.")
    print (("The program's file location is at %s") % BASE_DIR+  "\programs")
    print ("")
    #check for empty directory
    if len(os.listdir(BASE_DIR+"\programs")) < 1:
        print ("") 
        print("There are no files to list. Please record some.")
        time.sleep(2)
        main()
    items = os.listdir(BASE_DIR+"\\programs\\")
    fileList = [name for name in items if name.endswith(".wav")]
    for fcnt, fileName in enumerate(fileList, 1):
        sys.stdout.write("[%d] %s\n\r" % (fcnt, fileName))
    print ("")
    print ("[99] Main Menu")
    print ("")
    print ("[9999] Delete program")
    choice = (int(input("> ")))
   
    if (choice == 99):
        main()

    if (choice == 9999):
        os.system('cls')
        print ("Welcome to CoCo python tape player and recorder.")
        print ("For best results, please set PC sound volume between 54 to 58.")
        print (("The program's file location is at %s") % BASE_DIR+  "\programs")
        print ("")
        items = os.listdir("C:\\Users\\roger\\OneDrive\\Documents\\coco\\programs")
        for cnt, fileName in enumerate(fileList, 1):
            sys.stdout.write("[%d] %s\n\r" % (cnt, fileName))
        choice = int(input("Select WAV file to delete [1-%s]: " % cnt))-1
        fn = (fileList[choice])
        rf = (str("C:\\Users\\roger\\OneDrive\\Documents\\coco\\programs\\" + fn))
        #Delete file
        x = input("Are you sure you want to delete %s [Y/N]? >"% fn)
        if (x == "Y" or x == "y"):
            os.remove(rf)
            print ("%s file removed"% fn)
            time.sleep(1)            
            list() 
        else:
            list()
    if (choice < 90):
        list()       

def main():    
    #interface
    os.system('cls')
    print ("Welcome to CoCo python tape player and recorder.")
    print ("For best results, please set PC sound volume between 54 to 58.")
    print (("The program's file location is at %s") % BASE_DIR+  "\programs")
    print ("")
    print ("OPTIONS.")
    print ("")
    print ("[1] Play")
    print ("[2] Record")
    print ("[3] List Programs")
    print ("")
    print ("[99] Quit")
    x=input("Select a number option then press enter > ")

    os.system('cls')
    if (x == "1"):
        playtape()        
    if (x== "2"):
        record()
    if (x== "3"):
        list()
    if (x == "99"):
        quit()
    else:
        print ("INVALID OPTION. \n Please enter a valid option.")
        time.sleep (1.5)
        main()

os.system('cls')
main()


