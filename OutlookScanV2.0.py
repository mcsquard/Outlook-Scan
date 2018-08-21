# Python 3.6

import time
import datetime
import win32com.client
import cx_Oracle
import ParcelTools
import MLEntry
import StormTools
import os

INSPECTION_PATH = "F:\\ISGIS\\Investigations\\"   #"\\nas01p\\shared_dirs\\ccd\\gis\\gisdata\\ISGIS\\Investigations\\"
ML_Connection = MLEntry.ML_Connection
STORM_CONNECT_STR = "stmsvr/nilmd2s@OEGISP"
STM_Connection = cx_Oracle.connect(STORM_CONNECT_STR)
MLcursor = MLEntry.cursor
stmCursor = STM_Connection.cursor()
AuditorEmail = "Brooke.Sandoval@denvergov.org"

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
folder = outlook.Folders("PWWMD-CS Investigations")
inbox = folder.Folders("Inbox")
ForwardedToInvs = folder.Folders("Forwarded to Investigators")
ForwardedToTechs = folder.Folders("Forwarded To Techs")
EnteredInMasterlist = folder.Folders("Entered In Masterlist")

def GetParcelID(sched):
    stmCursor.execute("SELECT parcelid FROM parcel WHERE parcelnumber ='" + sched + "'")
    parcelid=stmCursor.fetchone()
    if parcelid is not None:
        return parcelid[0]
    else:
        return 0

def getInvInitials(email):
    Initials = MLcursor.execute("select INITIALS from INVESTIGATORS1 where EMAIL = '" + email + "'")
    return Initials.fetchone()[0]

def getTechInitials(email):
    Initials = MLcursor.execute("select INITIALS from CS_REPS1 where EMAIL = '" + email + "'")
    return Initials.fetchone()[0]

def duplicateCheck(schednum):
    qryStr = "SELECT ACCOUNT_NUMBER, REPRESENTATIVE, COMMENTS, REASON_ID, RECORD_ID FROM CMSTR WHERE ACCOUNT_NUMBER = '" + schednum + "' AND REP_COMPLETED_DATE is NULL"
    qry = MLcursor.execute(qryStr)
    dupcheck = qry.fetchone()
    if dupcheck is not None:
        #dup = dupcheck[1]
        return dupcheck
    else:
        return None

def getTechEmail(initials):
    sql = "select EMAIL from CS_REPS1 WHERE INITIALS='" + initials + "'"
    cur = MLcursor.execute(sql)
    email = cur.fetchone()
    if email is not None:
        return email[0]
    else:
        return None

# def checkforPrevious(sched):
#     previousIni = duplicateCheck(sched)[1]
#     if previousIni is not None:
#         techemail = getTechEmail(previousIni)
#         print(sched + " Previously Assigned to " + str(previousIni))
#         return techemail
#     else:
#         return None

def getNextTech():
    minQuery = "select MIN(ASSIGNMENTS) from CS_REPS1 where ACTIVE=1"
    minAssignments = MLcursor.execute(minQuery)
    nextTechQuery = "select * from CS_REPS1 where ACTIVE=1 and ASSIGNMENTS=" + str(minAssignments.fetchone()[0])
    nextTechCur = MLcursor.execute(nextTechQuery)
    nextTech = nextTechCur.fetchone()
    nextTechEmail = nextTech[5]
    return nextTechEmail

def AddTechAssignment(emailorinitials):
    querystr = "update CS_REPS1 set ASSIGNMENTS = ASSIGNMENTS + 1 where (EMAIL = '" + emailorinitials + "' or INITIALS = '" + emailorinitials + "')"
    MLcursor.execute(querystr)
    ML_Connection.commit()
    return

def SubTechAssignment(emailorintials):
    querystr = "select ASSIGNMENTS from CS_REPS1 where (EMAIL = '" + emailorintials + "' or INITIALS = '" + emailorintials + "')"
    res = MLcursor.execute(querystr)
    assigns = res.fetchone()[0]
    if assigns > 0:
        querystr = "update CS_REPS1 set ASSIGNMENTS = ASSIGNMENTS - 1 where (EMAIL = '" + emailorintials + "' or INITIALS = '" + emailorintials + "')"
        MLcursor.execute(querystr)
        ML_Connection.commit()
    else:
        querystr = "update CS_REPS1 set ASSIGNMENTS = 0 where (EMAIL = '" + emailorintials + "' or INITIALS = '" + emailorintials + "')"
        MLcursor.execute(querystr)
        ML_Connection.commit()
    return

def sendEmail(message):
    outlook = win32com.client.Dispatch("Outlook.Application")
    newMsg = outlook.CreateItem(0)
    newMsg.To = "matt.crowley@denvergov.org"
    newMsg.Subject = "OutlookScan.py"
    newMsg.body = message
    newMsg.send


def BuildInspPath(InspID):
    d = InspID[-8:] 
    schednum = InspID[:13]
    mapnum = schednum[:5]
    Path = INSPECTION_PATH + mapnum + "\\" + schednum + "\\" + schednum + " " + str(d[:4]) + "-" + str(d[4:6]) + "-" + str(d[-2:])
    return Path

class Employee:

    def __init__(self,emailaddr):
        self.firstname = emailaddr.split("@")[0].split(".")[0]
        self.lastname = emailaddr.split("@")[0].split(".")[1]
        self.email = emailaddr
        self.initials = self.getInitials(emailaddr)
        self.role = self.getRole(emailaddr)
    
    def FullName(self):
        return self.firstname + " " + self.lastname

    def getRole(self,email):
        qry = MLcursor.execute("select INITIALS from CS_REPS1 where EMAIL = '" + email + "'")
        results = qry.fetchone()
        if results is not None:
            return "TECH"
        else:
            qry = MLcursor.execute("select INITIALS from INVESTIGATORS1 where EMAIL = '" + email + "'") 
            results = qry.fetchone()
            if results is not None:
                return "INVESTIGATOR"
            else:
                return None

    def getInitials(self,email):
        qry = MLcursor.execute("select INITIALS from CS_REPS1 where EMAIL = '" + email + "'")
        results = qry.fetchone()
        if results is not None:
            return results[0]
        else:
            qry = MLcursor.execute("select INITIALS from INVESTIGATORS1 where EMAIL = '" + email + "'") 
            results = qry.fetchone()
            if results is not None:
                return results[0]
            else:
                return None

class Inspection(object):

    def __init__(self, EmailMsg):
        self.InspID = self.buildMessageID(EmailMsg)
        self.SchedNum = self.InspID[:13]
        self.CycleNum = StormTools.getCycleNumber(self.SchedNum[:5])
        self.InspDate = datetime.datetime(year=EmailMsg.SentOn.year,month=EmailMsg.SentOn.month,day=EmailMsg.SentOn.day)
        dup = duplicateCheck(self.SchedNum)
        if dup is not None:
            tech = Employee(getTechEmail(dup[1]))
            self.InspID = dup[2]
            self.InspType = dup[3]
        else:
            tech = Employee(getNextTech())
            self.InspType = 10
        self.Tech = tech
        self.Inv = Employee(EmailMsg.Sender.GetExchangeUser().PrimarySmtpAddress)
        self.Path = BuildInspPath(self.InspID)
    
    def buildMessageID(self, message): # MessageID is defined as the attachment name without spaces or the .pdf extension
        #Determine if an open record exists in ML. 
        attachname = message.Attachments[0]
        sched = str(attachname)[:13]
        rec=MLEntry.findOpenRecord(sched)
        #If so, use that InspID
        if rec is not None and rec[3].isnumeric == True:
            messageID = rec[3]
            return messageID
        # If not create one from file name.
        else:
            splitname = str(attachname).split(" ")
            messageID = splitname[0] + splitname[1]
            messageID = messageID.split(".")[0]
            return messageID
    
    def getTechEmail(self,initials):
        sql = "select EMAIL from CS_REPS1 WHERE INITIALS='" + initials + "'"
        cur = MLcursor.execute(sql)
        email = cur.fetchone()
        if email is not None:
            return email[0]
        else:
            return None
    
    def ForwardInspection(self,msg):
        newMsg = msg.Forward()
        newMsg.To = self.Tech.email
        msgbody = "<HTML><BODY>Reply <a href='mailto:pwwmd.investigations@denvergov.org?subject=Re%3A " + self.SchedNum + " - Completed&body=Hit Send and do not edit anything in this message. Thank you.%0A%0AInvestigator: " + self.Inv.email + "%0AInvestigation ID: " + self.InspID + "'>Completed</a> when done."
        msgbody = msgbody + "<BR>Project folder: <a href='" + self.Path + "'>" + self.Path + "</a>"
        newMsg.HTMLBody = msgbody
        if self.InspType == 2 or self.InspType == 7:
            newMsg.Importance = 2
        newMsg.Send()
        print(self.InspID + " Forwarded to " + self.Tech.email)

class Completed(object):

    def __init__(self, EmailMsg):
        self.InspID = self.getInspID(EmailMsg)
        self.SchedNum = self.InspID[:13]
        self.AssignDate = datetime.datetime(year=EmailMsg.SentOn.year,month=EmailMsg.SentOn.month,day=EmailMsg.SentOn.day)
        self.CompleteDate = datetime.datetime(year=EmailMsg.SentOn.year,month=EmailMsg.SentOn.month,day=EmailMsg.SentOn.day)
        self.Tech = Employee(EmailMsg.Sender.GetExchangeUser().PrimarySmtpAddress) 
    

    def getInspID(self, email):
        Lines = email.body.splitlines()
        #lastLine = str(Lines[-1]).replace(" ", "")
        try:
            lastLine = str(Lines[6]).replace(" ", "")
        except IndexError:
            lastLine = str(Lines[-1]).replace(" ", "")
        ID = lastLine.split(":")[-1]
        return ID
    
    def CompleteinML(self):
        MLEntry.UpdateMLfield(self.InspID, "REPRESENTATIVE", self.Tech.initials)
        MLEntry.UpdateMLfield(self.InspID, "REP_COMPLETED_DATE", self.CompleteDate)
        Items = ForwardedToTechs.Items
        n = Items.Count
        for n in range(n,0,-1):
            message = Items.Item(n)
            newInsp=Inspection(message)
            if newInsp.InspID == self.InspID:
                message.move(EnteredInMasterlist)
        Items = ForwardedToInvs.Items
        n = Items.Count
        for n in range(n,0,-1):
            message = Items.Item(n)
            req=InspectionRequest(message)
            if req.SchedNum == self.SchedNum:
                message.move(EnteredInMasterlist)
        print(self.InspID + " Completed in Masterlist by " + self.Tech.email)

      
class Audit(Completed):
    def __init__(self,msg):
        Completed.__init__(msg)

    def CompleteAudit(self):
        MLEntry.UpdateMLfield(self.InspID,"AUDITOR","BMS")
        MLEntry.UpdateMLfield(self.InspID,"AUD_COMPLETED_DATE",datetime.datetime.today())

class InspectionRequest(object):
    
    def __init__(self,msg):
        #Parse message body
        splitbody = msg.body.splitlines()
        splitbody = list(filter(None,splitbody))
        sched=splitbody[0].split("=")[1]
        sched = sched.replace("-","")
        sched = sched.replace(" ","")
        insptype = splitbody[1].split("=")[1]
        insptype = insptype.strip()
        comment = splitbody[2].split("=")[1]
        comment = comment.strip() 
        # for n in range(3,len(splitbody)-1):
        #     comment = comment + str(splitbody[n][0])
        #     comment = comment.strip()
        if len(sched)==13 and sched.isnumeric()==True:
        #Check for valid parcel/storm account?
            parcel = ParcelTools.ParcelBySchedNum(sched)
            if parcel is not None:
                self.SchedNum = parcel.schednum
                self.InspID = self.SchedNum + str(msg.SentOn.year) + str(msg.SentOn.month) + str(msg.SentOn.day)   #str(datetime.date.today()) #should be date sent
                self.InspID = self.InspID.replace("-","")
                self.InitiateDate = datetime.datetime(year=msg.SentOn.year,month=msg.SentOn.month,day=msg.SentOn.day)   #should be date sent
                self.InvAssignDate = datetime.date.today()
                self.TechAssignDate = ""
                self.Tech = Employee(msg.Sender.GetExchangeUser().PrimarySmtpAddress)
                self.Inv = self.getInv("",self.SchedNum[:5])
                self.InspType = insptype
                self.InspPath = self.getPath(self.SchedNum)
                self.InspComment = comment
            else:
                self.SchedNum = None

    def getInv(self,email="", mapnum=""):
        if len(mapnum)==5 and mapnum.isnumeric()==True:
            qry = MLcursor.execute("SELECT investigator FROM mappings WHERE map = '" + mapnum + "'")
            ini = qry.fetchone()[0]
            qry = MLcursor.execute("select email from INVESTIGATORS1 where initials = '" + ini + "'")
            e=qry.fetchone()[0]
            return Employee(e)
        else:
            return Employee(email)
        
    def getTechInitials(self,email):
        result = MLcursor.execute("select INITIALS from CS_REPS1 where EMAIL = '" + email + "'")
        Initials = result.fetchone()
        if Initials is not None:
            return Initials[0]
        else:
            return None
    def getInvEmail(self, email):
        Lines = email.body.splitlines()
        try:
            line = str(Lines[4]).replace(" ", "")
        except IndexError:
            line = str(Lines[-2]).replace(" ", "")
        InvEmail = line.split(":")[-1]
        return InvEmail

    def getPath(self, sched):
        newInspPath = INSPECTION_PATH + sched[:5] + "\\" + sched + "\\" + sched + " " + str(datetime.date.today())
        return newInspPath

    def ForwardRequest(self,msg):
        newMsg = msg.Forward()
        newMsg.Subject = self.InspType + ": " + self.SchedNum
        newMsg.To = self.Inv.email
        newMsg.HTMLbody = "<html><body>New Inspection: <a href='" + self.InspPath + "'>" + self.InspPath + "</a>"
        if self.InspType == "Customer Request" or self.InspType == "Title Call":
            newMsg.Importance = 2 #High Importance
        newMsg.Send()
        print(self.InspID + " Forwarded to " + self.Inv.email)
        msg.move(ForwardedToInvs)

class Process(object):       

    def __init__(self):
        #First go through all emails in the inbox and separate them into new Investigations and completed Investigations
        NewMessages = inbox.Items
        n = NewMessages.Count
        #If inbox has zero messages just get out.
        if n==0:
            return
        #Process each message
        for n in range(n,0,-1):
            try:
            #Determine sender is tech or investigator
                message = NewMessages.Item(n)
                sender = Employee(message.Sender.GetExchangeUser().PrimarySmtpAddress)
                if message.subject[:2].upper() == "RE" and sender.role == "TECH":
                    #Completed
                    self.CompleteInvestigation(message)
                elif message.Subject == "New Inspection Request" and sender.role == "TECH":
                    #New Inspection
                    self.InitiateInvestigation(message)
                elif sender.role == "INVESTIGATOR" and message.Attachments.Count > 0:
                    #Storm Update or Report from previous initiation
                    self.ProcessInvestigation(message)
                elif sender == AuditorEmail:
                    a=Audit(message)
                    a.CompleteAudit()
                else:
                    self.ReplyToSender(message)
            except:
                message.move("Quarantine")
                continue

    def InitiateInvestigation(self,msg):
        newInsp = InspectionRequest(msg)
        #if invalid kick it back to tech, otherwise
        if newInsp.SchedNum is not None:
            if not os.path.exists(newInsp.InspPath):
                os.makedirs(newInsp.InspPath)
            for att in msg.Attachments:
                Filename = newInsp.InspPath + "\\" + att.Filename
                att.SaveAsFile(Filename)
            #Enter In Masterlist
            parcel = ParcelTools.ParcelBySchedNum(newInsp.SchedNum)
            #MLEntry.ML_TABLE = "CMSTR1" #Change this upon GoLive
            dup = duplicateCheck(newInsp.SchedNum)
            if dup is None:
                MLEntry.PutInMasterList(parcel.schednum, parcel.cycleNum, newInsp.InspID, parcel.addNum, parcel.addPrefix, parcel.addStreet, parcel.addSuffix, parcel.addUnit, newInsp.InitiateDate, newInsp.Inv.initials, newInsp.InitiateDate, "", newInsp.Tech.initials, "","",newInsp.InspType)
            else:
                MLEntry.UpdateMLfield(dup[4],"COMMENTS", newInsp.InspID)
                MLEntry.UpdateMLfield(dup[4],"INVESTIGATOR",newInsp.Inv.initials)
                MLEntry.UpdateMLfield(dup[4],"INV_ASSIGNED_DATE",newInsp.InvAssignDate)
                MLEntry.UpdateMLfield(dup[4],"REPRESENTATIVE",newInsp.Tech.initials)
            #AddTechAssignment(newInsp.Tech.email)
        #Forward Inspection Request to Investigator with path to folder
            newInsp.ForwardRequest(msg)
            #msg.move(ForwardedToInvs)
        else:
            self.ReplyToSender(msg) #return the email
        
    def ProcessInvestigation(self,msg):
        newInsp = Inspection(msg)
        if not os.path.exists(newInsp.Path):
            os.makedirs(newInsp.Path)
        for att in msg.Attachments:
            Filename = newInsp.Path + "\\" + att.Filename
            att.SaveAsFile(Filename)
        parcel = ParcelTools.ParcelBySchedNum(newInsp.SchedNum)
        if parcel is None:
            masterAcct = StormTools.getStormAcct2(newInsp.SchedNum)
            altSched = StormTools.getParcelFromAcct(masterAcct)
            parcel = ParcelTools.ParcelBySchedNum(altSched)
            parcel.schednum = newInsp.SchedNum
        if parcel is not None:
            dup = duplicateCheck(newInsp.SchedNum)
            if dup is None:
                MLEntry.PutInMasterList(newInsp.SchedNum, parcel.cycleNum, newInsp.InspID, parcel.addNum, parcel.addPrefix, parcel.addStreet, parcel.addSuffix, parcel.addUnit, newInsp.InspDate, newInsp.Inv.initials, newInsp.InspDate, newInsp.InspDate, newInsp.Tech.initials, "","","Storm Update")
                AddTechAssignment(newInsp.Tech.initials)
            else:
                MLEntry.UpdateMLfield(newInsp.InspID, "INV_COMPLETED_DATE", newInsp.InspDate)
            newInsp.ForwardInspection(msg)
            msg.move(ForwardedToTechs)
        else:
            self.ReplyToSender(msg)

    def CompleteInvestigation(self, msg):
        compMsg = Completed(msg)
        compMsg.CompleteinML()
        msg.move(EnteredInMasterlist)
        cur = MLEntry.cursor
        sql = "SELECT reason_id FROM " + MLEntry.ML_TABLE +  " WHERE COMMENTS = '" + compMsg.InspID + "'"
        results = cur.execute(sql)
        result = results.fetchone()
        if result is not None:
            reason_id = result[0]
            if reason_id != 10:
                newMsg = msg.Forward()
                newMsg.To = AuditorEmail
                newMsg.Subject = msg.Subject
                path = BuildInspPath(compMsg.InspID)
                msgbody = "<HTML><BODY>Reply <a href='mailto:pwwmd.investigations@denvergov.org?subject=Re%3A " + compMsg.SchedNum + " - Completed&body=Hit Send and do not edit anything in this message. Thank you.%0A%0AInvestigator: " + compMsg.Tech.initials + "%0AInvestigation ID: " + compMsg.InspID + "'>Completed</a> when done."
                msgbody = msgbody + "<BR>Project folder: <a href='" + path + "'>" + path + "</a>"
                #newMsg.HTMLbody = "Completed Inspection: <b>" + compMsg.SchedNum + "</b><br>Inspection Directory: <a href='" + path + "'>" + path + "</a>" 
                newMsg.HTML = msgbody
                newMsg.Send()
                print(compMsg.InspID + " Forwarded to " + AuditorEmail + " for audit")
            else:
                SubTechAssignment(compMsg.Tech.email)




    def CheckStorm(self):
        Items = ForwardedToTechs.Items
        n = Items.Count
        for n in range(n,0,-1):
            message = Items.Item(n)
            newInsp=Inspection(message)
            parcelid=GetParcelID(newInsp.SchedNum)
            if parcelid != 0:
                sql="SELECT ENTRYDATE FROM parceldetail WHERE (reason='NEW DETAIL' AND parcelid=" + str(parcelid) + ") ORDER BY ENTRYDATE DESC"#+ " AND measurementdate >= to_date('" + str(newInsp.InspDate) + "','mm/dd/yyyy hh:mi:ssam'))"
                stmCursor.execute(sql)
                details = stmCursor.fetchall()
                #insp=datetime.datetime.fromtimestamp(int(str(newInsp.InspDate)))
                for detail in details:
                    newInsp.InspDate=newInsp.InspDate.replace(tzinfo=None)
                    entrydate = detail[0]
                    if entrydate >= newInsp.InspDate:
                        if newInsp.Tech is not None:
                            MLEntry.UpdateMLfield(newInsp.InspID,"REPRESENTATIVE", newInsp.Tech.initials)
                            MLEntry.UpdateMLfield(newInsp.InspID,"REP_COMPLETED_DATE", entrydate)
                            print(newInsp.InspID + " Completed in Storm by " + newInsp.Tech.email)
                            SubTechAssignment(newInsp.Tech.email)
                            message.move(EnteredInMasterlist)
                    break

    def ReplyToSender(self,msg):
        newMsg = msg.Reply()
        newMsg.Subject = "error: " + msg.Subject
        newMsg.To = msg.Sender.GetExchangeUser().PrimarySmtpAddress
        newMsg.body = "There was an Error encountered in the processessing of this Inspection Request.\nIt has not been forwarded to an investigator or entered in Masterlist.\nCheck for a valid parcel number."
        newMsg.body = newMsg.body + msg.body
        print("Error email sent to " + str(msg.Sender.GetExchangeUser().PrimarySmtpAddress))
        newMsg.Send()
    





timesrun = 0
#p.CheckStorm()
while True:
    try:
        p=Process()
        if timesrun == 16:
            sendEmail("Success" + str(time.ctime()))
            runonce = False
            timesrun=0
        print (time.ctime())
        timesrun = timesrun + 1
        p = None
        time.sleep(900)
        
    except Exception as e:
        sendEmail("Error-ProcessInbox()" + "\n" + str(e))
        print("Error-ProcessInbox()- " + str(e))
        p = None
        time.sleep(900)
        continue
        
  
