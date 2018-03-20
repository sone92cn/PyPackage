import threading

class ReportMessage():
    def __init__(self, mid=0, code=0, msg=""):
        self.mid = mid
        self.code = code
        self.msg = msg
        """
        mid: int mission id if used
        code: int messagecode 0:error else:user defined
        msg: str string info for msg
        """
    def __repr__(self):
        return "mid=" + str(self.mid) + ",code=" + str(self.code) + ",msg="  + self.msg
        
class MissionEx():
    def __init__(self, v_list, v_inx, func=None, args=(), submit_self=False, report_call=None, finished=True):
        self.func = func
        self.args = args
        self.v_list = []
        self.v_inx = v_inx
        self.finished = finished
        self.lock = threading.Lock()
        self.events = []
        self.valid = True
        self.msg = []
        self.success = 0
        self.report_call = report_call
        self.submit_self = submit_self
        if len(v_list):
            self.v_list.extend(v_list)
            if finished == True:
                if self.v_list[len(self.v_list)-1] != None:
                    self.v_list.append(None)
            elif self.v_list[len(self.v_list)-1] == None:
                finished=True
        else:
            finished=False
    def createMission(self, func, args=()):
        self.func = func
        self.args = args
        return self
    def resetVarList(self, v_list, finished=True):
        self.lock.acquire()
        if len(v_list):
            if finished == True:
                if v_list[len(v_list)-1] != None:
                    v_list.append(None)
            elif v_list[len(v_list)-1] == None:
                finished=True
        else:
            finished=False
        self.v_list = v_list
        self.valid = True
        self.msg = []
        self.success = False
        self.lock.release()
    def addMission(self, v_obj):
        self.lock.acquire()
        if self.v_list == None:
            self.lock.release()
            return False
        if self.finished:
            self.lock.release()
            return False
        self.v_list.append(v_obj)
        if len(self.events):
            if v_obj == None:
                self.finished = True
                while (len(self.events)):
                    event = self.events.pop()
                    event.set()
            else:
                event = self.events.pop()
                event.set()
        self.lock.release()
        return True
    def extendMission(self, v_list):
        self.lock.acquire()
        if self.v_list == []:
            self.lock.release()
            return False
        if self.finished:
            self.lock.release()
            return False
        self.v_list.extend(v_list)
        if len(self.events):
            if v_list == []:
                self.finished = True
                while (len(self.events)):
                    event = self.events.pop()
                    event.set()
            else:
                event = self.events.pop()
                event.set()
        self.lock.release()
    def endMission(self):
        self.lock.acquire()
        if self.v_list == []:
            self.v_list.append(None)
            self.finished = True
        elif self.v_list[len(self.v_list)-1] != None:
            self.v_list.append(None)
            self.finished = True
        if len(self.events):
            while (len(self.events)):
                event = self.events.pop()
                event.set()
        self.lock.release()
        return True
    def getMission(self, event=None):
        self.lock.acquire()
        if (self.func == None) or (len(self.v_list) == 0):
            if event == None:
                self.lock.release()
                return None
            else:
                self.events.append(event)
                self.lock.release()
                event.wait()
                mission = self.getMission()
                while (mission == None):
                    self.lock.acquire()
                    self.events.append(event)
                    self.lock.release()
                    event.wait()
                    mission = self.getMission()
                return mission
        else:
            if self.v_list[0] == None:
                self.lock.release()
                self.valid = False
                return (0, None, None)  #code=1 exit()
            else:
                arg = self.v_list.pop(0)
                args = []
                args.extend(self.args[:self.v_inx])
                args.append(arg)
                args.extend(self.args[self.v_inx+1:])
                if self.submit_self:
                    args.append(self)
                self.lock.release()
                return (1, self.func, tuple(args))
    def addReport(self, msg=None):
        if self.report_call != None and msg != None:
            self.msg.append(msg)
            self.success = msg.code
            self.report_call(msg.mid, msg.code, msg.msg)

class ManagerEx():
    def __init__(self, mission, nthreads=4, block_primary=True, new_thread=False, end_call=None, end_args=()):
        """end_call will run with <error_code, end_args> when missions finished, only if self.block_primary is true"""
        self.nrunning = 0
        self.mission = mission
        self.nthreads = nthreads
        self.new_thread = new_thread
        self.end_call = end_call
        self.end_args = end_args
        self.lock = threading.Lock()
        self.event = threading.Event()
        self.block_primary = block_primary
    def __start(self):
        if self.nrunning == 0:
            self.lock.acquire()
            n = min(self.nthreads, len(self.mission.v_list))
            while (self.nrunning < n):
                thread = threading.Thread(target=self.__run, args=(threading.Event(),))
                thread.setDaemon(True)
                thread.start()
                self.nrunning += 1
            self.lock.release()
            if self.block_primary:
                self.event.wait()
                if self.end_call != None:
                    self.end_call(self.mission.success, self.mission.msg, self.end_args)
    def __run(self, event):
        code, func, args = self.mission.getMission(event)
        while code:
            if len(args):
                msg = func(args)
            else:
                msg = func()
            self.mission.addReport(msg)
            code, func, args = self.mission.getMission(event)
        self.lock.acquire()
        self.nrunning -= 1
        if self.nrunning == 0:
            self.event.set()
        self.lock.release()
    def start(self):  
        """return self.error if new_thread=False and bolck_primry=True
        else return True"""
        if self.new_thread:
            thread = threading.Thread(target=self.__start)
            thread.setDaemon(True)
            thread.start()
        else:
            self.__start()
            if self.block_primary:
                return self.mission.success
        return True