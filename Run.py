import pyautogui
import mysql.connector
import speech_recognition as sr
import time
import re
import sys
import os
import win32com.client
import random
import webbrowser

path="'C:\Program Files\Google\Chrome\Application\chrome.exe' %s"#Enter file location of the default browser for this application
google="http://www.google.com/search?q="
speak = win32com.client.Dispatch('Sapi.SpVoice')
speak.Rate=0.0001
speak.Volume=100
try:
    r = sr.Recognizer()
    mic=sr.Microphone()
    commands=['move to position (x,y)','move up by 100','move down by 100','move right by 100','move left by 100','right click','left click','middle click','double click','type','press','show commands','drag to position(x,y)','drag up by 100','drag down by 100','drag right by 100','drag left by 100','open calculator','search','quit','shutdown','restart']
    keys=['-','=',',',"'",'.','tab', 'enter', 'space', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', 'accept', 'add', 'alt', 'altleft', 'altright', 'apps', 'backspace', 'browserback', 'browserfavorites', 'browserforward', 'browserhome', 'browserrefresh', 'browsersearch', 'browserstop', 'capslock', 'clear', 'convert', 'ctrl', 'ctrlleft', 'ctrlright', 'decimal', 'del', 'delete', 'divide', 'down', 'end', 'enter', 'esc', 'escape', 'execute', 'f1', 'f10', 'f11', 'f12', 'f13', 'f14', 'f15', 'f16', 'f17', 'f18', 'f19', 'f2', 'f20', 'f21', 'f22', 'f23', 'f24', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9', 'final', 'fn', 'hanguel', 'hangul', 'hanja', 'help', 'home', 'insert', 'junja', 'kana', 'kanji', 'launchapp1', 'launchapp2', 'launchmail', 'launchmediaselect', 'left', 'modechange', 'multiply', 'nexttrack', 'nonconvert', 'num0', 'num1', 'num2', 'num3', 'num4', 'num5', 'num6', 'num7', 'num8', 'num9', 'numlock', 'pagedown', 'pageup', 'pause', 'pgdn', 'pgup', 'playpause', 'prevtrack', 'print', 'printscreen', 'prntscrn', 'prtsc', 'prtscr', 'return', 'right', 'scrolllock', 'select', 'separator', 'shift', 'shiftleft', 'shiftright', 'sleep', 'space', 'stop', 'subtract', 'tab', 'up', 'volumedown', 'volumemute', 'volumeup', 'win', 'winleft', 'winright', 'yen', 'command', 'option', 'optionleft', 'optionright']
    conn=mysql.connector.connect(host='localhost',user='root',passwd='1234')#change hostname,usernam,password according to your mysql
    cur=conn.cursor()
    try:
        cur.execute('create database commands')
        conn.commit()
    except:
        pass
    conn=mysql.connector.connect(host='localhost',user='root',passwd='1234',database='commands')#change hostname,usernam,password according to your mysql
    cur=conn.cursor()
    try:
        cur.execute('create table commands(name varchar(100), value int(10))')
        conn.commit()
        for i in commands:
            cur.execute('insert into commands value("'+i+'",0)')
            conn.commit()
            
    except:
        pass

    def text2int (text):
        if ',' in text:
            text=text.replace(',',' comma')
        words=text.split()
        ans=words[:]
        units = ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight","nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen","sixteen", "seventeen", "eighteen", "nineteen"]
        tens = ["" ,"", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"]
        scales=["hundred","thousand"]
        for i in range (len(words)):
            if words[i] in units+tens+scales+['and']:
                res=0
                nums_in_words=[]
                while i<len(words) and words[i] in units+tens+scales+['and']:
                    nums_in_words.append(words[i])
                    i+=1
                if 'thousand' in nums_in_words:
                    res+=1000
                    remaining=nums_in_words[nums_in_words.index('thousand'):]
                if 'hundred' in nums_in_words:
                    if nums_in_words.index('hundred')!=0:
                        for j in range(nums_in_words.index('hundred')):
                            if nums_in_words[j] in units:
                                res+=units.index(nums_in_words[j])*100
                            elif nums_in_words[j] in tens:
                                res+=tens.index(nums_in_words[j])*1000
                        remaining=nums_in_words[nums_in_words.index('hundred'):]
                    else:
                        res+=100
                        remaining=nums_in_words[nums_in_words.index('hundred'):]
                if 'hundred' not in nums_in_words and 'thousand' not in nums_in_words:
                    remaining=nums_in_words[:]
                for j in remaining:
                    if j in units:
                        res+=units.index(j)
                    elif j in tens:
                        res+=tens.index(j)*10
                insert_index=ans.index(nums_in_words[0])
                for j in nums_in_words:
                    ans.remove(j)
                ans.insert(insert_index,str(res))

                #if it contains more than 1 numerical values in the form of words
                c=[i for i in ans if i in units or i in tens or i in scales]
                if len(c)>0:
                    return text2int(" ".join(ans))
                else:
                    return " ".join(ans)
        return text

    def display_commands():
        print('This is the list of available commands')
        speak.Speak('This is the list of available commands')
        commands_=['move to position (x,y); maximum value:'+str(pyautogui.size())+'your present mouse location is:'+str(pyautogui.position()),'move up(specify the amount like move up by 100)','move down(move down by 100)','move right(move right by 100)','move left(move left by 100)','drag to position (x,y); maximum value:'+str(pyautogui.size())+'you present mouse location is:'+str(pyautogui.position()),'drag up(specify the amount  drag up by 100)','drag down(drag down by 100)','drag right(drag right by 100)','drag left(drag left by 100)','right click','left click','middle click','double click','type( type python)','press (after this the name of keys like press ctrl a)','show commands','open (starts any file in your computer)','search(google serch)','quit(it will end the program)','shutdown','restart']
        num=1
        for num in commands_:
            print(num,i)
            num+=1
            speak.Speak(i)
        
        print("Chose one(only 1 request is processed at a time) of the commands(out of these predefined ones) to execute, I would recommend to speak exact same commands in same order (part in brackets is not necessary).")
        speak.Speak('Chose one(only 1 request is processed at a time) of the commands(out of these predefined ones) to execute, I would recommend to speak exact same commands in same order (part in brackets is not necessary). I am waiting for 5 seconds before starting.')
        time.sleep(5)
        main()

    def execute(instruction):
        if 'move' in instruction:
            instruction=text2int(instruction)
            num=list(map(int,re.findall(r'\d+',instruction)))
            if 'up' in instruction:
                pyautogui.moveRel(0, -(num[0]), duration=1)
                cur.execute('update commands set value=value+1 where name="move up by 100"')
                conn.commit()
            elif 'position' in instruction:
                x,y=num[0],num[1]
                pyautogui.moveTo(x,y,1)
                cur.execute('update commands set value=value+1 where name="move to position (x,y)"')
                conn.commit()
            elif 'down' in instruction:
                pyautogui.moveRel(0, (num[0]), duration=1)
                cur.execute('update commands set value=value+1 where name="move down by 100"')
                conn.commit()
            elif 'right' in instruction:
                pyautogui.moveRel(num[0], 0, duration=1)
                cur.execute('update commands set value=value+1 where name="move right by 100"')
                conn.commit()
            elif 'left' in instruction:
                pyautogui.moveRel(-(num[0]), 0, duration=1)
                cur.execute('update commands set value=value+1 where name="move left by 100"')
                conn.commit()
            else:
                print('Please recheck the recognition result(if it matches with what you spoke),try to give commands more similar to predefined ones. Sorry, I am not able to detect it.')
                speak.Speak('Please recheck the recognition result(if it matches with what you spoke),try to give commands more similar to predefined ones. Sorry, I am not able to detect it.')
                main()
        elif 'drag' in instruction:
            instruction=text2int(instruction)
            num=list(map(int,re.findall(r'\d+',instruction)))
            if 'up' in instruction:
                pyautogui.dragRel(0, -(num[0]), duration=1)
                cur.execute('update commands set value=value+1 where name="drag up by 100"')
                conn.commit()
            elif 'position' in instruction:
                x,y=num[0],num[1]
                pyautogui.dragTo(x,y,1)
                cur.execute('update commands set value=value+1 where name="drag to position (x,y)"')
                conn.commit()
            elif 'down' in instruction:
                pyautogui.dragRel(0, (num[0]), duration=1)
                cur.execute('update commands set value=value+1 where name="drag down by 100"')
                conn.commit()
            elif 'right' in instruction:
                pyautogui.dragRel(num[0], 0, duration=1)
                cur.execute('update commands set value=value+1 where name="drag right by 100"')
                conn.commit()
            elif 'left' in instruction:
                pyautogui.dragRel(-(num[0]), 0, duration=1)
                cur.execute('update commands set value=value+1 where name="drag left by 100"')
                conn.commit()
            else:
                print('Please recheck the recognition result(if it matches with what you spoke),try to give commands more similar to predefined ones. Sorry, I am not able to detect it.')
                speak.Speak('Please recheck the recognition result(if it matches with what you spoke),try to give commands more similar to predefined ones. Sorry, I am not able to detect it.')
                main()
        elif 'click' in instruction:
            if 'right' in instruction:
                pyautogui.rightClick()
                cur.execute('update commands set value=value+1 where name="right click"')
                conn.commit()
            elif 'left' in instruction:
                pyautogui.leftClick()
                cur.execute('update commands set value=value+1 where name="left click"')
                conn.commit()
            elif 'middle' in instruction:
                pyautogui.middleClick()
                cur.execute('update commands set value=value+1 where name="middle click"')
                conn.commit()
            elif 'double' in instruction:
                pyautogui.doubleClick()
                cur.execute('update commands set value=value+1 where name="double click"')
                conn.commit()
            else:
                print('Please recheck the recognition result(if it matches with what you spoke),try to give commands more similar to predefined ones. Sorry, I am not able to detect it.')
                speak.Speak('Please recheck the recognition result(if it matches with what you spoke),try to give commands more similar to predefined ones. Sorry, I am not able to detect it.')
                main()
        elif 'type' in instruction:
            pyautogui.typewrite(instruction.split('type')[-1])
            cur.execute('update commands set value=value+1 where name="type"')
            conn.commit()
        elif 'command' in instruction:
            display_commands()
            cur.execute('update commands set value=value+1 where name="show commands"')
            conn.commit()
        elif 'quit' in instruction:
            print('Thanks for using this program')
            speak.Speak('Thanks for using this program')
            time.sleep(2)
            cur.execute('update commands set value=value+1 where name="quit"')
            conn.commit()
            os._exit(1)
        elif 'press' in instruction:
            cur.execute('update commands set value=value+1 where name="press"')
            conn.commit()
            instruction=instruction.split()
            key_names={'spacebar':'space','control':'ctrl','function':'fn','minus':'-','equal':'=',"plus":"+",'comma':",",'quote':"'",'dot':".",'pause':"."}
            for i in range(len(instruction)):
                if instruction[i] in key_names:
                    instruction[i]=key_names[instruction[i]]
            instruction=[i for i in instruction if i in keys]
            for i in instruction:
                pyautogui.keyDown(i)
            for j in instruction:
                pyautogui.keyUp(i)
        elif 'shut' in instruction:
            speak.Speak('Shutting down')
            cur.execute('update commands set value=value+1 where name="shutdown"')
            conn.commit()
            cur.close()
            conn.close()
            os.system('shutdown /s /t 1')
            exit()
        elif 'restart' in instruction:
            speak.Speak('Restarting')
            cur.execute('update commands set value=value+1 where name="restart"')
            conn.commit()
            cur.close()
            conn.close()
            os.system('shutdown /r /t 1')
            exit()
        elif 'search' in instruction:
            cur.execute("update commands set value=value+1 where name='search'")
            conn.commit()
            instruction_array=instruction.split()
            instruction_array=instruction_array[instruction_array.index('search')+1:]
            search_query="+".join(instruction_array)
            run=google+search_query
            webbrowser.get(path).open(run)
        elif 'open' in instruction:
            cur.execute('update commands set value=value+1 where name="open calculator"')
            conn.commit()
            pyautogui.press('win')
            instruction_array=instruction.split()
            instruction_array=instruction_array[instruction_array.index('open')+1:]
            pyautogui.typewrite(' '.join(instruction_array))
            pyautogui.press('enter')
        else:
            print("We haven't received any command similar to the command list, please try again")
            speak.Speak("We haven't received any command similar to the command list, please try again")
            main()
    
    def main():
        with mic as source:
            print('Listening')
            speak.Speak('I am listening')
            audio= r.record(source=mic, duration=10)
            speak.Speak('Wait till I execute')
            print('Executing')
            
        try:
            recognition_text=r.recognize_google(audio, language='en-IN').lower()
            print("You said " + recognition_text)
            speak.Speak('You said'+recognition_text)
            execute(recognition_text)
            speak.Speak('Command executed')
            cur.execute('select name from commands order by value asc')
            fetched_commands=cur.fetchall()[0:9]
            print('Try using these commands:',end="")
            for _ in range(3):
                random_command=random.choice(fetched_commands)
                print(random_command[0],end=", ")
                fetched_commands.remove(random_command)
            print()
            main()
        except:
            speak.Speak("I did't get any value. Start speaking when I am listening.")
            main()

    cur.execute('select name from commands order by value asc')
    fetched_commands=cur.fetchall()[0:9]
    print('Try using these commands:',end="")
    speak.Speak('Try using these commands')
    for _ in range(3):
        random_command=random.choice(fetched_commands)
        print(random_command[0],end=", ")
        speak.Speak(random_command[0])
        fetched_commands.remove(random_command)
    print()
    print('for curated list of all commands use the command "show commands"')
    speak.Speak('for curated list of commands use the command "show commands"')
    main()
except:
    print('Microphone should be connected to run this program')
    speak.Speak('Please check if your microphone is connected')