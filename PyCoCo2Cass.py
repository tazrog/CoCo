#Written by tazrog
#Designed to save programs on the PC via cassette interface and load them in the CoCo.
#Program still and develpment.
#Use on your own risk. No guarantees.
#For use on a Windows PC. Not tested on Linux or Mac.

from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import wave
import pyaudio
from array import array
import os
import sys
import keyboard
import time
from tqdm import tqdm
global frame_rate
global auto_sound
global auto_sound_level
frame_rate = 9600
auto_sound =1
auto_sound_level = -9.0
BASE_DIR = os.getcwd()
if not os.path.exists(BASE_DIR+"\\Programs"):
    os.makedirs(BASE_DIR+"\\Programs") 

def sound():
    global soundlevel
    global volume
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)    
    volume = interface.QueryInterface(IAudioEndpointVolume)
    if (auto_sound == 0):
        soundlevel = volume.GetMasterVolumeLevelScalar() * 100
    #Set PC Sound to 55 if auto=sound is on.
    if (auto_sound == 1):
        soundlevel= volume.GetMasterVolumeLevel()   
        volume.SetMasterVolumeLevel(auto_sound_level, None)        

def record():
    os.system('cls')
    print ("Welcome to CoCo Python tape player and recorder on PC.")
    print (("The program's file location is at %s") % BASE_DIR+  "\\Programs")
    print ("")
    print ("RECORDING SECTION")
    print ("")
    fcnt = 0
    FORMAT=pyaudio.paInt16
    CHANNELS=1
    RATE=frame_rate
    CHUNK=1024
    RECORD_SECONDS=1
    wavitems = os.listdir(BASE_DIR+"\\Programs")
    wavList = [wav for wav in wavitems if wav.endswith(".wav")]    
    for fcnt, wavName in enumerate(wavList, 1):
        sys.stdout.write("[%d] %s\n\r" % (fcnt, wavName))      
    if len(os.listdir(BASE_DIR+"\\Programs")) < 1:
        tapename = 9999 
        #print ("[9999}New File Name")
        #print ("")         
    else:
        print ("")
        print ("[99] Main Menu")
        print ("[9999] New File Name")        
        print ("")
    print ("Type in your option [number] or a file [number] of a file you will like to override.")
    print ("")
    while True:
        if (len(os.listdir(BASE_DIR+"\\Programs")) > 0):
            tapename = (input(">>> "))
        try:
            tapename = int(tapename)
        except ValueError:
            os.system('cls')
            print ("INVALID SELECTION")
            time.sleep(1)
            record()        
        if (tapename == 99):
            main()            
        if (tapename == 9999):
            newname = input("Type new file name. Do not include .wav extension. >>> ")
            rfn = newname +".wav"        
        if (tapename > fcnt and tapename != 99):
            if (tapename != 9999):
                os.system('cls')
                print ("INVALID SELECTION")
                time.sleep(1)
                record() 
        else:
            tapename = tapename -1
            rfn = (wavList[tapename])
            option = str(input ("Are you sure you wany to override file "+ rfn +" [y/n]? >>> "))   
            if (option != "y"):
                print ("File override canceled")
                time.sleep(1)
                record()        
        selection = input ("Enter the CSAVE command on the CoCo and press any key on the PC to start recording. \nAfter PC is recording execute CSAVE on PC. Enter [x] to exit. ")
        if (selection == "x"):
            print ("Program Exited")
            main()
        FILE_NAME=(BASE_DIR+"\\Programs\\" + rfn)
        os.system('cls')
        print ("RECORDING...")        
        print ("")
        print (("The program's file location is at %s") % BASE_DIR+  "\\Programs")
        print ("***********************************************************************************************")
        print (rfn)
        print ("***RECORDING NOW*** PRESS [q] to stop recording when CoCo is done saving.")
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
                    time.sleep (3)
                    os.system('cls')
                    keyboard.send('backspace')
                    main()   
                data=stream.read(CHUNK)
                frames.append(data)   

def playtape():
    os.system('cls') 
    sound()
    print ("Welcome to CoCo Python tape player and recorder on PC.")
    print (("The program's file location is at %s") % BASE_DIR+  "\\Programs")
    print ("")
    print ("PLAYING TAPE SECTION")

    #check for empty directory
    if len(os.listdir(BASE_DIR+"\\Programs")) < 1: 
        print ("")
        print("There are no files to play. Please record some.")
        volume.SetMasterVolumeLevel(soundlevel, None)
        time.sleep(2)
        main()
    chunk = 1024      
    fcnt =0 
    items = os.listdir(BASE_DIR+"\\Programs")
    fileList = [name for name in items if name.endswith(".wav")]
    for fcnt, fileName in enumerate(fileList, 1):
        sys.stdout.write("[%d] %s\n\r" % (fcnt, fileName))
    print ("")
    print ("[99] Main Menu")
    print ("")    
    choice= (input("Select WAV file[1-%s] or main menu [99] option: " % fcnt))    
    while True: 
        try:            
            choice = int(choice)-1
            if (choice > 95):
                main()            
        except ValueError:
            os.system('cls')
            print ("INVALID SELECTION")
            time.sleep(1)
            playtape()  
        if (choice+1 > fcnt):
                os.system('cls')
                print ("INVALID SELECTION")
                time.sleep(1)
                playtape()  
        fn = (fileList[choice])
        wavplay = BASE_DIR + "\\Programs\\"+ fn
        f = wave.open(wavplay,"rb")
        p = pyaudio.PyAudio()        
        stream = p.open(format = p.get_format_from_width(f.getsampwidth()),  
            channels = f.getnchannels(),  
            rate = f.getframerate(),  
            output = True)
        frames = f.getnframes()
        rate = f.getframerate()
        duration = (round(frames/rate))
        mins, secs = divmod(duration, 60) 
        os.system('cls')
        print ("File to be played:")
        print (wavplay) 
        print ("Frame Rate: %d"% f.getframerate())  
        print ("Frame Channel: %d"% f.getnchannels())  
        print ('Duration = {:02d}:{:02d}'.format(mins, secs))
        if (auto_sound == 1):
                print ("Auto-sound is On")
                if (auto_sound_level == -9.0):
                    print ("Auto Sound level is at 55 percent.")
                if (auto_sound_level == -10.5):
                    print ("Auto Sound level is at 50 percent.")
                if (auto_sound_level == -7.6):
                    print ("Auto Sound level is at 60 percent.") 
        if (auto_sound == 0):
            sound()
            print ("Auto-sound is off")
            print ("PC audio level is set to %d percent"% (soundlevel + 1))
            print ("**********************************************************************************************")
        print ("Execute CLOAD on the CoCo and press enter when ready on the PC.")
        print ("Or enter [99] Main Menu")        
        print ("")
        selection = (input("[Enter] to play Wav once CLOAD is executed on the CoCo >>> ")) 
        if (selection == "99"):                 
            stream.stop_stream()  
            stream.close() 
            p.terminate()
            f.close()
            main()                   
        print (("Playing file >>> %s")% (fileList[choice]))
        print ("")
        data = f.readframes(chunk)        
        countdown= (round(frames/chunk)) + 1

        #play stream  
        for i in tqdm(range(countdown),desc = "Load Progress:"):             
            stream.write(data)  
            data = f.readframes(chunk)        
  
        #stop stream        
        stream.stop_stream()  
        stream.close()  

        #close PyAudio 
        
        p.terminate()
        f.close()  
        print ("Tape Complete")   
        if (auto_sound == 1):
            volume.SetMasterVolumeLevel(soundlevel, None)
        main()      

def list():
    os.system('cls')
    print ("Welcome to PyCoCo2Cass tape player and recorder on PC.")
    print (("The program's file location is at %s") % BASE_DIR+  "\\Programs")
    print ("")
    fcnt = 0    
    #check for empty directory
    if len(os.listdir(BASE_DIR+"\\Programs")) < 1:
        print ("") 
        print("There are no files to list. Please record some.")
        time.sleep(2)
        main()
    items = os.listdir(BASE_DIR+"\\Programs")
    fileList = [name for name in items if name.endswith(".wav")]
    for fcnt, fileName in enumerate(fileList, 1):
        sys.stdout.write("[%d] %s\n\r" % (fcnt, fileName))    
    print ("")
    print ("[99] Main Menu")
    print ("[9999] Delete program")
    print ("")
    while True:
        choice = (input("Select WAV file [1-%s] to get Wav details. >>> "% fcnt))        
        try:
            choice = int(choice)-1
            if (choice+1 == 99):
                main()
            if (choice < 9990):
                if (choice > fcnt):
                    print ("INVALID SELECTION")
                    time.sleep(1)
                    list()
        except ValueError:
            os.system('cls')
            print ("INVALID SELECTION")
            time.sleep(1)
            list()
        if (choice < fcnt):
            fn = (fileList[choice])
            wavplay = BASE_DIR + "\\Programs\\"+ fn
            f = wave.open(wavplay,"rb")
            p = pyaudio.PyAudio()        
            stream = p.open(format = p.get_format_from_width(f.getsampwidth()),  
                channels = f.getnchannels(),  
                rate = f.getframerate(),  
                output = True)
            frames = f.getnframes()
            rate = f.getframerate()
            duration = (round(frames/rate))
            mins, secs = divmod(duration, 60)
            os.system('cls')
            print ("***File Information***")
            print ("File Location %s"% wavplay) 
            print ("")
            print ("File name: %s"% fn)
            print ("Frame Rate: %d"% f.getframerate())  
            print ("Frame Channel: %d"% f.getnchannels())  
            print ('Duration = {:02d}:{:02d}'.format(mins, secs)) 
            print ("")
            a = input ("Press enter to return to list menu. ")
            list()
              
        else:                    
            os.system('cls')
            print ("Welcome to CoCo python tape player and recorder.")
            print (("The program's file location is at %s") % BASE_DIR+  "\\Programs")
            print ("")
            items = os.listdir(BASE_DIR + "\\Programs")
            dcnt=0
            for dcnt, fileName in enumerate(fileList, 1):
                sys.stdout.write("[%d] %s\n\r" % (dcnt, fileName))
            while True:
                dchoice = (input("Select WAV file to delete [1-%s]: " % dcnt))
                try:
                    dchoice = int(dchoice)-1
                    if (dchoice > dcnt and dchoice < 9999):
                        print ("INVALID SELECTION")
                        time.sleep(1)
                        list()
                except ValueError:
                    os.system('cls')
                    print ("INVALID SELECTION")
                    time.sleep(1)
                    list()
                fn = (fileList[dchoice])
                rf = str(BASE_DIR + "\\Programs\\" + fn)
                #Delete file
                x = input("Are you sure you want to delete %s [Y/N]? >"% fn)
                if (x == "Y" or x == "y"):
                    os.remove(rf)
                    print ("%s file removed"% fn)
                    time.sleep(1)            
                    main() 
                else:               
                    os.system('cls')
                    print ("INVALID SELECTION")
                    time.sleep(1)
                    list()                    

def settings():
    #in beta
    global frame_rate
    global auto_sound
    global auto_sound_level
    os.system('cls')
    print ("Settings")
    print ("OPTIONS")   
    print ("[1] Frame Rate")
    print ("[2] Sound")
    print ("")
    print ("[99] Main Menu")
    setinput=int(input(">>> "))    
    if (setinput == 1):
        os.system('cls')
        print ("Current Recording Frame Rate \(Baud\) is : %d" % frame_rate)
        print ("Frame Rate Options:")
        print ("[1] 1500")
        print ("[2] 9600")
        print ("[3] 44100")
        print ("[4] 48000")
        
        print ("")
        b= input(">>> ")
        if (b== "1"):
            frame_rate = 1500
        if (b == "2"):
            frame_rate = 9600
        if (b=="3"):
            frame_rate = 44100
        if (b == "4"):
            frame_rate = 48000
        
        frame_rate = int(frame_rate)
        print ("")
        print ("New fram rate is %d: "% frame_rate)
        c = input ("Press any key to go to main menu")
        main()
    elif (setinput == 2):
        os.system('cls')
        if (auto_sound == 1):
            print ("Auto-sound is On")
        else:
            print ("Auto-sound is off")
        if (auto_sound_level == -9.0):
            print ("Auto Sound level is at 55 percent.")
        if (auto_sound_level == -10.5):
            print ("Auto Sound level is at 50 percent.")
        if (auto_sound_level == -7.6):
            print ("Auto Sound level is at 60 percent.")
        print ("")
        print ("[1] Turn on auto-sound.")
        print ("[2] Turn off auto-sound.")
        print ("[3] Adjust auto-sound setting to 50 percent.")
        print ("[4] Adjust auto-sound setting to 55 percent *recommeded.")
        print ("[5] Adjust auto-sound setting to 60 percent.")
        print ("")
        print ("Any other key to go to Main Menu")
        print ("")
        a = (input (">>> "))
        if (a == "1"):
            auto_sound = 1
            auto_sound = int(auto_sound)
            print (" Auto sound is on")
            c = input("Press any key to go to setting menu. ")
            settings()
        if ( a == "2"):
            auto_sound = 0
            auto_sound = int(auto_sound)
            print (" Auto sound is off")
            c = input("Press any key to go to setting menu. ")
            settings()
        if (a =="3"):
            auto_sound_level = -10.5
            auto_sound_level = float(auto_sound_level)
            print ("Sound level set to 50 percent.")
            c = input("Press any key to go to setting menu. ")
            settings()
        if (a == "4"):
            auto_sound_level = -9.0
            auto_sound_level = float(auto_sound_level)
            print ("Sound level set to 55 percent.")
            c = input("Press any key to go to setting menu. ")
            settings()
        if (a == "5"):
            auto_sound_level = -7.6
            auto_sound_level = float(auto_sound_level)
            print ("Sound level set to 60 percent.")
            c = input("Press any key to go to setting menu. ")
            settings()
       
        else:
            main()

    else:
        main()

def main():    
    #interface 
   
    os.system('cls')
    print ("Welcome to CoCo Python tape player and recorder on PC.")    
    print (("The program's file location is at %s") % BASE_DIR+  "\\Programs")
    print ("")
    print ("OPTIONS.")
    print ("")    
    print ("[1] Play")
    print ("[2] Record")
    print ("[3] List/Delete Programs")
    print ("[0] Settings")
    print ("")
    print ("[99] Quit")
    print ("")
    while True:
        x=input("Select a number option, then press enter >>> ")
        try:
            x =int(x)
        except ValueError:
            os.system('cls')
            print ("INVALID OPTION [%s]"% x)
            time.sleep(1)
            main()
        if (x == 0):
            settings()
        if (x == 1):
            playtape()        
        if (x == 2):
            record()
        if (x == 3):
            list()
        if (x == 99):
            quit()
        else:
            os.system('cls')
            print ("INVALID OPTION")
            print ("Not a valid option [%d]."% x)
            time.sleep (1)
            main()

#Program Start
os.system('cls')
main()