from flask import Flask, render_template, request,send_file
from flask import Flask, send_from_directory
import pandas as pd
import random
classes=[]
app = Flask(__name__)
TOTAL_HRS=7
DAYS=5
MAX_SIZE=120
GAP=17

def populate(s):
    length=len(s)
    list_1=[]
    class_ind=[]
    for k in range(length):
        if str(s.iat[k,1]).lower()!='nan':
            class_ind.append(list((s.iat[k,0],s.iat[k,1])))
        else:
            if  class_ind:
                list_1.append(class_ind)
            classes.append(str(s.iat[k,0]))
            class_ind=[]
    list_1.append(class_ind)
    return list_1

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/page2.html')
def page2():
    return render_template('page2.html')

@app.route('/downloads/<path:filename>')
def download_file(filename):
    return send_from_directory('static', filename, as_attachment=True)


@app.route('/view', methods=['POST'])
def view():
    files = [request.files[f'file{i}'] for i in range(1, 4)]
   # file1.save(file1.filename)
    s2=pd.read_excel(files[0])
    s1=pd.read_excel(files[1])
    s3=pd.read_excel(files[2])
    #s1=pd.read_excel("Course_teacher_map.xlsx")



    teachers=s1['faculty'].dropna().unique().tolist()
    print(teachers)

    teacher_course=populate(s1)
    print(teacher_course)
    teacher_len=len(teachers)

    #s2=pd.read_excel("ti.xlsx")
    t_len=len(s2)
    print(t_len)
    #s3=pd.read_excel("course_hour_map.xlsx")
    course_hour=populate(s3)
    print(course_hour)
    timeslot=[[0]*MAX_SIZE for i in range(t_len)]
    teacherslot=[[0]*MAX_SIZE for i in range(teacher_len)]

    for i in range(t_len):
        index=0
        for k in range(DAYS):
            for j in range(TOTAL_HRS):
                timeslot[i][index]=str(s2.iat[i,(k*TOTAL_HRS)+j])
                index+=1
            index+=GAP


    for k in range(t_len):
        c_t=teacher_course[k]
        c_h=course_hour[k]
        for i in range(len(c_h)):
            course=c_h[i][0]
            hour=c_h[i][1]
            fac=str(s1.loc[s1["course"]==course,'faculty'].item())
            t_index=teachers.index(fac)
            alloc_hr=0
            rem_hr=hour
            slots=[]
            while rem_hr>0:
                for j in range(MAX_SIZE):
                    if str(timeslot[k][j]).lower()=="nan":
                        begin=j
                        break
                for j in range((MAX_SIZE-1),-1,-1):
                    if str(timeslot[k][j]).lower()=="nan":
                        end=j
                        break
                if rem_hr!=1:
                    interval=(end-begin+1)/(rem_hr-1)
                pos=begin
                for j in range(int(rem_hr)):
                    slots.append(pos)
                    pos=pos+interval
                for slot in slots:
                    if str(timeslot[k][int(slot)]).lower()=="nan" and teacherslot[t_index][int(slot)]==0:
                        timeslot[k][int(slot)]=course
                        teacherslot[t_index][int(slot)]=course
                    else:
                        left=int(slot)-1
                        right=int(slot)+1
                        while left>0 or right<MAX_SIZE:
                            if((left>0 and str(timeslot[k][left]).lower()=="nan")) and teacherslot[t_index][left]==0:
                                timeslot[k][left]=course
                                teacherslot[t_index][left]=course
                                break
                            if(right<MAX_SIZE and str(timeslot[k][right]).lower()=="nan") and teacherslot[t_index][right]==0:
                                timeslot[k][right]=course
                                teacherslot[t_index][right]=course
                                break
                            left=left-1
                            right=right+1
                        if left<0 and right>=MAX_SIZE:
                            print("ERROR: ALLOCATION COULD NOT BE DONE")
                            break
                            
                    rem_hr-=1
                else:
                    continue
                break       
            
    k=0
    timetable=[]
    while k<t_len:
        index=0
        temp=[[0]*TOTAL_HRS for i in range(DAYS)]
        timetable.append(['','','',classes[k],'','',''])
        for i in range(DAYS):
            for j in range(TOTAL_HRS):
                if str(timeslot[k][index]).lower()=="nan":
                    timeslot[k][index] ="REMEDIAL"
                temp[i][j]=timeslot[k][index]           
                index+=1
            timetable.append(temp[i])
            index+=GAP
        
        k=k+1
        
    print(timetable)
    tf=pd.DataFrame(timetable,index=['','monday','tuesday','wednesday','thursday','friday']*t_len,columns=['1st','2nd','3rd','Lunch','4th','5th','6th'])
    tf.to_excel("static/tf7.xlsx")
    return send_file("static/tf7.xlsx", as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
                    