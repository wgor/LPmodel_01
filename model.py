import xlwings as xw
from mesa import *
from mesa.time import BaseScheduler
import pandas as pd
import numpy as np
from pulp import *
from math import floor

# Excelfile import
excelfile = ("input_file.xlsm")
wb = xw.Book(excelfile)

# clear costs from previous runs
for sh in range(1,len(wb.sheets)):
    wb.sheets[sh].range("A15:A16").value = 0

class EnergyModel(Model):
    # model with some entities
    def __init__(self):

        # list of named activated agents by sheetname
        active_agents = ["a"+str(int(i)) for i in wb.sheets['IO'].range("I3:I12").value if int(i) > 0]
        print ("Active Agents: {}".format(active_agents))
        self.num_agents = len(active_agents)
        self.schedule =  BaseScheduler(self)

        # imports timesteps=[1:96] and datetimes(e.g. 00:00)
        self.time = wb.sheets['a1'].range("C3:D99").options(pd.Series).value

        # timesteps=[1:96]
        self.timeindex = self.time.index

        # Create Prosumer agents
        for i in active_agents:
            name = i
            a = Prosumer(i,self, name)
            self.schedule.add(a)

    def step(self):
        self.schedule.step()


class Prosumer(Agent):
    # agent with pv, battery and demand
    def __init__(self, unique_id, model,name):
        super().__init__(unique_id, model)

        # agents Properties
        self.name = name
        self.costs = 0
        self.run_status = 0

        # agent's Excelsheet
        self.sht = wb.sheets[self.name]
        self.agent_ts = self.sht.range("C3:Q99").options(pd.DataFrame).value

        # get agent parameter input
        self.agent_p = self.sht.range("A1:B13").options(pd.Series).value
        self.paramdict = self.agent_p.to_dict()

        # get dict values for each agent
        self.min_dis = self.paramdict.get("min_dis")
        self.max_dis = self.paramdict.get("max_dis")
        self.min_char = self.paramdict.get("min_cha")
        self.max_char = self.paramdict.get("max_cha")
        self.thres_down = self.paramdict.get("thres_down")
        self.thres_up = self.paramdict.get("thres_up")
        self.batt_eff = self.paramdict.get("batt_eff")
        self.max_buy = self.paramdict.get("max_buy")
        self.max_sell = self.paramdict.get("max_sell")
        self.SOC_init = self.paramdict.get("initSOC")
        self.SOC_end = self.paramdict.get("endSOC")
        self.horizont = self.paramdict.get("horizont")

        # PV and DEMAND TIMESERIES Input
        self.pv = self.agent_ts.loc[:,"pv"]
        self.dem = self.agent_ts.loc[:,"dem"]

        # MARKET TIMESERIES
        self.mp = self.agent_ts.loc[:,"mp"]
        self.fp = self.agent_ts.loc[:,"fp"]

    def step(self):
        self.optimize()
        self.writeToXL()
        pass

    def optimize(self):
        # creates periods (-> input agent's HORIZONT)
        self.timeframes = self.periodIndexer()
        # LP Relaxation for all timeframes
        self.lpRelax()
        return

    def writeToXL(self):
        wb = xw.Book(excelfile)
        sht = wb.sheets[self.name]
        sht.range('C3:Q99').value = self.agent_ts
        sht.range('Total_Costs').value = value(self.costs)
        sht.range('Status').value = self.run_status

    def periodIndexer(self):
        index = len(self.agent_ts.index)
        period_length = self.horizont
        period_overlay=index%period_length
        fullperiods = floor(index/period_length)
        periodinits = []

        for p in range(1,fullperiods+1):
            periodinits.append(period_length * p)
        if period_overlay > 0:
            periodinits.append(index)
        return periodinits

    def lpRelax(self):

        firstindex = 0
        lastindex = 0
        lastcap = 0

        # 'prob' variable
        lpmodel = LpProblem("AGENT_OPT_01",LpMinimize)
        print ("Run: {}".format(lpmodel.name))

        # call every period in self.timeframes and optimize it one by one
        for i,j in enumerate(self.timeframes):
            print("Periodnr.: {}".format(i+1))
            actindex = int(j)
            if i == 0:
                runtime = self.agent_ts.index[firstindex:actindex]
                print("Timesteps: 0-{}".format(str(actindex)))
            else:
                runtime = self.agent_ts.index[lastindex:actindex]
                print ("Timesteps: {}".format(str(lastindex) + "-" + str(actindex)))
            lastindex = actindex

            # TIMESERIES VARIABLES
            buy = LpVariable.dicts("buy", runtime, 0,upBound=self.max_buy, cat= "NonNegativeReals")
            sell = LpVariable.dicts("sell", runtime, 0,upBound=self.max_sell, cat= "NonNegativeReals")
            cap = LpVariable.dicts("batt_cap", runtime, lowBound=self.thres_down,upBound=self.thres_up, cat= "NonNegativeReals")
            dis = LpVariable.dicts("discharged", runtime, cat= "NonNegativeReals")
            char = LpVariable.dicts("charged", runtime, cat= "NonNegativeReals")
            b_stat = LpVariable.dicts("stat_buy", runtime, cat= "Binary")
            s_stat = LpVariable.dicts("stat_sell", runtime, cat= "Binary")
            d_stat = LpVariable.dicts("stat_dis", runtime, cat= "Binary")
            c_stat = LpVariable.dicts("stat_char", runtime, cat= "Binary")

            # OBJECTIVE
            lpmodel += lpSum(buy[t]*self.mp[t]-sell[t]*self.fp[t] for t in runtime)

            # BATTERY CONSTRAINT 1: wether charging or discharging in t
            for t in runtime:
                lpmodel += d_stat[t] + c_stat[t] <= 1

            # BATTERY CONSTRAINT 2: char and dischar limits, if stat is 1
            for t in runtime:
                lpmodel += self.min_dis*d_stat[t] <= dis[t]
                lpmodel += self.max_dis*d_stat[t] >= dis[t]
                lpmodel += self.min_char*c_stat[t] <= char[t]
                lpmodel += self.max_char*c_stat[t] >= char[t]

            # BATTERY CONSTRAINT 4: battery cap must be between low and high threshold
            for t in runtime:
                lpmodel += cap[t] >= self.thres_down
                lpmodel += cap[t] <= self.thres_up

            # BATTERY CONSTRAINT 5: Init and End State of BatteryCap
            for t in runtime:
                past = t-1
                if t == 1:
                    lpmodel += cap[1] == self.SOC_init
                    lpmodel += dis[1] == 0
                    lpmodel += char[1] == 0

                elif t == int(min(runtime)):
                    lpmodel += cap[int(min(runtime))] == lastcap
                    lpmodel += dis[int(min(runtime))] == 0
                    lpmodel += char[int(min(runtime))] == 0

                else:
                    lpmodel += cap[t]==cap[past]-dis[t]+char[t]

                if t == max(self.model.timeindex):
                    lpmodel += cap[t] == self.SOC_end

            # BALANCING CONSTRAINT: sold + pv + dis == bought + demand + char per step
            for t in runtime:
                lpmodel += buy[t]+self.pv[t]+dis[t] == sell[t]+self.dem[t]+char[t]

            # MARKET CONSTRAINT 1: maximum buy and sell quantity per step
            for t in runtime:
                lpmodel += buy[t] <= self.max_buy*b_stat[t]
                lpmodel += sell[t] <= self.max_sell*s_stat[t]

            # MARKET CONSTRAINT 2: buying and selling in same step is not possible => Arbitrage
            for t in runtime:
                lpmodel += b_stat[t] + s_stat[t] <= 1

            ## SOLVE MODEL FOR All PERIODS
            #LpSolverDefault.msg = 1
            lpmodel.solve()
            self.costs += value(lpmodel.objective)
            self.run_status = LpStatus[lpmodel.status]
            print ("Agentname: {}, Costs: {}, Model_Status: {} \n".format(self.name,round(self.costs), self.run_status))

            ## WRITE OUTPUT DATA TO AGENT's TS
            for t in runtime:
                self.agent_ts.loc[t,"sell"] = sell[t].varValue*-1
                self.agent_ts.loc[t,"buy"] = buy[t].varValue
                self.agent_ts.loc[t,"cap"] = cap[t].varValue
                self.agent_ts.loc[t,"b_stat"] = b_stat[t].varValue
                self.agent_ts.loc[t,"s_stat"] = s_stat[t].varValue
                self.agent_ts.loc[t,"c_stat"] = c_stat[t].varValue
                self.agent_ts.loc[t,"d_stat"] = d_stat[t].varValue
                self.agent_ts.loc[t,"char"] = char[t].varValue
                self.agent_ts.loc[t,"dis"] = dis[t].varValue*-1
             ## save actual batt cap for next period
            lastcap = cap[max(runtime)].varValue
        return
