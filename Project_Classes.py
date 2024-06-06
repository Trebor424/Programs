# -*- coding: utf-8 -*-

#moje klasy

class Wyrob:
    def __init__(self, masterid, Panel_Serial_number, boardsnumber ,parts ,aux ,value ,loc ,el ,reference ,tolerance_plus ,tolerance_minus ,timestamp,testvalue,proces):
        self.masterid=masterid
        self.Panel_Serial_number=Panel_Serial_number
        self.boardsnumber=boardsnumber
        self.parts=parts
        self.aux=aux
        self.value=value
        self.loc=loc
        self.el=el
        self.reference=reference
        self.tolerance_plus=tolerance_plus
        self.tolerance_minus=tolerance_minus
        self.testvalue=testvalue
        self.proces=proces
        self.timestamp=timestamp
        
        
class Testy_wyrobow:
    def __init__(self, parts,aux,value,boardsnumber, loc, el, reference, tolerance_plus, tolerance_minus, testvalue, timestamp):
        self.parts=parts
        self.aux = aux
        self.value = value
        self.boardsnumber=boardsnumber
        self.loc = loc
        self.el = el
        self.reference = (reference)
        self.tolerance_plus = (tolerance_plus)
        self.tolerance_minus = (tolerance_minus)
        self.listtestvalue = [(testvalue), (timestamp)]

class Wyniki_testow:
    def __init__(self,fixture, testname ,group, cp , cpk , pp , ppk , o_stand,limitchange):
        self.fixture=fixture,
        self.testname=testname,
        self.group=group,
        self.cp=cp,
        self.cpk=cpk,
        self.pp=pp,
        self.ppk=ppk,
        self.o_stand=o_stand
        self.limitchange=limitchange
    
    
