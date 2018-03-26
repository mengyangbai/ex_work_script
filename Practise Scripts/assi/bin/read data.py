# @author
import os
import collections
import copy

data_dir = r"D:\workspace\pythonScripts\python\Practise Scripts\assi\data"
filename = "EXPL4.DAT"
output_serial = r"D:\workspace\pythonScripts\python\Practise Scripts\assi\output_serial"
output_parallel = r"D:\workspace\pythonScripts\python\Practise Scripts\assi\output_parallel"



'''
    project : one project
        graph_dict : process graph
        days_cost : each process cost
        resource_request : each resource request
        max_resources : resource you have
'''


class project(object):
    

    def get_all_paths(self,all_jobs,order_in,graph_dict,reverse_graph):
        if not all_jobs:
            self.projects.append(order_in)
            return self.projects
        
        for element in all_jobs:
            jobs = copy.deepcopy(all_jobs)
            order = copy.deepcopy(order_in)
            jobs.remove(element)
            order.append(element)
            for succ in graph_dict[element]:
                if succ not in jobs and succ not in order:
                    for pred in reverse_graph[succ]:
                        if pred not in order:
                            break
                    else:
                        jobs.append(succ)
        
            self.get_all_paths(jobs, order, graph_dict, reverse_graph)
            
        return self.projects

    '''
        time : the current time that is to observed
    '''
    def available_activity(self, scheduled_set, time, start_times, require_list):
        check_set = []
        for activity in require_list:
            if int(activity) in scheduled_set:
                continue
                
            check = True
            # To check that the previous word is all finished
            for pred in self.reverse_dict[activity]:
                pred = int(pred)
                if pred not in start_times:
                    check = False
                    
            if check:
                check = True
                for pred in self.reverse_dict[activity]:
                    pred = int(pred)
                    activity = int(activity)
                    # To check, at the observed time the previous job is finished
                    if start_times[pred] + self.days_cost[pred] > time:
                        check = False
                        break
                if check:
                    check_set.append(activity)

        return check_set

    def parallel(self,require_list):
        start_times = {1:0}
        resources_spent = [[],[],[],[]]
        scheduled_set = [1]
        observation_time = 0
        while len(scheduled_set) != self.numbers: #stopping condition of recursion
            check = False
            available_set = self.available_activity(scheduled_set, observation_time, start_times, require_list)
            for activity in available_set:

                for time in range(observation_time, observation_time + self.days_cost[activity]):
                    for i in range(len(self.max_resources)):
                        while len(resources_spent[i]) <= time:
                            #补零
                            resources_spent[i].append(0)

                        if resources_spent[i][time] + int(self.resource_request[activity][i]) > self.max_resources[i]:#set resource for each resource type with i
                            check=True
                            break
                    if check:
                        break

                if not check:
                    scheduled_set.append(activity)
                    start_times[activity] = observation_time
                    for time in range(observation_time, observation_time + self.days_cost[activity]):
                        for i in range(len(self.max_resources)):
                            resources_spent[i][time] = resources_spent[i][time] + int(self.resource_request[activity][i]) 
            
            observation_time += 1
                
        
        return start_times, start_times[12]


    '''
        Serial algorithm
    '''
    def serial(self,require_list):
        '''
            Start_times : n th jobs : job start time
            resource_spent : the job spent on each time
            sheduled : whether it is sheduled or not
            check : to confirm the current resource is available to be putted here or not
        '''
        start_times = {1:0}
        resources_spent = [[],[],[],[]]

        while require_list:
            activity = int(require_list.pop(0)) # pop next activity in schedule
            start_time= 0

            # caculate the earliest time to begin
            for pred in self.reverse_dict[str(activity)]:
                pred = int(pred)
                if start_time <  start_times[pred] + self.days_cost[pred]:
                    start_time =  start_times[pred] +  self.days_cost[pred]

            scheduled = False
            check = False

            while not scheduled:
                
                if self.days_cost[activity] == 0:
                    start_times[activity] = start_time
                    scheduled = True

                else:
                    start_times[activity] = start_time
                    
                    # check the resouce is fitted in current positiion 
                    # if it is not jump out and add time
                    for time in range(start_time, start_time + self.days_cost[activity]):
                        for i in range(len(self.max_resources)):
                            while len(resources_spent[i]) <= time:
                                #补零
                                resources_spent[i].append(0)

                            if resources_spent[i][time] + int(self.resource_request[activity][i]) > self.max_resources[i]:#set resource for each resource type with i
                                check=True
                                break
                        if check:
                            break
                    
                    if not check:
                        for time in range(start_time, start_time + self.days_cost[activity]):
                            for i in range(len(self.max_resources)):
                                while len(resources_spent[i]) <= time:
                                    resources_spent[i].append(0)
                                resources_spent[i][time] = resources_spent[i][time] + int(self.resource_request[activity][i]) 
                        
                        scheduled = True
                    else:
                        check = False

                start_time = start_time + 1
        
        return start_times, start_times[self.numbers]



    def __init__(self, filename):
        
        self.projects=[]
        
        input_file = os.path.join(data_dir, filename)
        
        with open(input_file, 'r') as f:
            file_lines = list(f.readlines())
            
        self.graph_dict = {}
        self.days_cost = {}
        self.resource_request = {}

        self.max_resources = []

        graph_start = 0
        graph_end = 0

        resource_start = 0
        resource_end = 0

        available_resource_index = 0

        for i in range(len(file_lines)):

            if "jobnr.    #modes  #successors   successors" in file_lines[i]:
                graph_start = i + 1

            if "REQUESTS/DURATIONS:" in file_lines[i]:
                graph_end = i - 1

            if "------------------------------------------------------------------------" in file_lines[i]:
                resource_start = i + 1

            if "RESOURCEAVAILABILITIES:" in file_lines[i]:
                resource_end = i - 1
                available_resource_index = i + 2

        graph_lines = file_lines[graph_start:graph_end]
        self.numbers = len(graph_lines)
        duration_lines = file_lines[resource_start:resource_end]
        available_resource_line = file_lines[available_resource_index]

        for tmp_line in graph_lines:
            tmp_line = tmp_line.lstrip("   ")
            tmp_line = tmp_line.rstrip("\n")
            tmp_list = tmp_line.split()
            if len(tmp_list) == 3:
                self.graph_dict[str(tmp_list[0])] = set()
            else:
                self.graph_dict[str(tmp_list[0])] = set(tmp_list[3:])

        self.reverse_dict={}
        self.reverse_dict["1"]=set()
        for key in self.graph_dict:
            for value in self.graph_dict[key]:
                if value in self.reverse_dict:
                    self.reverse_dict[value].add(key)
                else:
                    self.reverse_dict[value] = set()
                    self.reverse_dict[value].add(key)

        
        # for key in self.graph_dict:
        #     print(key+":"+str(self.graph_dict[key]))

        # for key in self.reverse_dict:
        #     print(key+":"+str(self.reverse_dict[key]))


        for tmp_line in duration_lines:
            tmp_line = tmp_line.lstrip("   ")
            tmp_line = tmp_line.rstrip("\n")
            tmp_list = tmp_line.split()
            self.days_cost[int(tmp_list[0])] = int(tmp_list[2])
            self.resource_request[int(tmp_list[0])] = tmp_list[3:]

        self.max_resources = list(map(int,available_resource_line.lstrip(
            "   ").rstrip("\n").split()))

        

        self.resources=[]
        for i in range(self.numbers):
            self.resources.append(self.max_resources)
            


        self.projects = self.get_all_paths(["1"],[],self.graph_dict,self.reverse_dict)


        serial_file = os.path.join(output_serial, filename)
        
        parallel_file = os.path.join(output_parallel, filename)

        with open(parallel_file,'w') as f:
            tmp_projects = copy.deepcopy(self.projects)
            total_results = 0
            number = len(tmp_projects)
            
            for one_route in tmp_projects:
                sequence, result=self.parallel(one_route)
                total_results = total_results + result
                f.write(str(sequence)+"\t"+str(result)+"\n")
        
            f.write("The average parallel algorithm is : "+str(total_results/ float(number)))
        

        with open(serial_file, 'w') as f:
            tmp_projects = copy.deepcopy(self.projects)
            total_results = 0
            number = len(tmp_projects)
            for one_route in tmp_projects:
                sequence, result=self.serial(one_route)
                total_results = total_results + result
                f.write(str(sequence)+"\t"+str(result)+"\n")
            
            f.write("The average serial algorithm is : "+str(total_results/ float(number)))

        # print(projects[0])
        # print(self.parallel(projects[0]))



      





if __name__ == "__main__":
    # project(filename)
    for root, dirs, files in os.walk(data_dir):
        for single_file in files:
            if single_file.endswith(".DAT"):
                print("Operating file : "+ single_file)
                project(single_file)

    print("Work done!")
    # print(test.all_jobs)
    # get_all_paths(["1"],[],test.graph_dict,test.reverse_dict)
    # print(",".join(test.result_lists))
    #test.serial(tmplist)

