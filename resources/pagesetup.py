#!/usr/bin/env python
# -*- coding: UTF-8 -*-
#
import wx
import wx.grid
import os
import subprocess
from collections import defaultdict
import ConfigParser
import copy
try:
    from reportlab.lib.pagesizes import landscape, portrait
    from reportlab.pdfgen import canvas
except :
    import os, sys
    pp = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    sys.path.insert(0, pp)
    from reportlab.lib.pagesizes import landscape, portrait
    from reportlab.pdfgen import canvas
import rmsprt.num2word as num2word  
import rmsprt.raw_constant as rw
import rmss_config
from rmstxt.text_validator import  Rmss_TextCtrl_Num, MyValidator, FLOAT_ONLY
from rmss_head import RMS_ToolTip

"""
import re, win32api
import win32print
def ffile(folder, rex):
    for r, dirs, files in os.walk(folder):
        for f in files :
            result = rex.search(f)
            if result:
                return os.path.join(r, f)
    return None

def ffalldrive(drive=None, prymfolder=None, searchfolder=None, filename=None):
    rex = re.compile(filename)
    getipath = None
    for p in os.listdir(drive):
        if prymfolder in p:
            getipath =  os.path.join(drive, p, searchfolder)
    #for d in win32api.GetLogicalDriveStrings().split('\000'):
    #    print d
    return ffile(getipath, rex)
"""



def Get_pgsett_Path(resc='resources', pg1='pgsett.ini', pg2='pgcompsett.ini', pg3='pgsettA4.ini'):
    PGPATH1 = '\\'.join([resc, pg1])
    PGPATH2 = '\\'.join([resc, pg2])
    PGPATH3 = '\\'.join([resc, pg3])
    return PGPATH1, PGPATH2, PGPATH3

class PRINT_DISCRIPTION(object):
    def __init__ (self):
        #self.stfile = rmss_config.my_icon('resources\\pgsett.ini')
        self.PGPATH1, self.PGPATH2, self.PGPATH3 = Get_pgsett_Path()
        self.stfile = rmss_config.my_icon(self.PGPATH1)
        configg = ConfigParser.ConfigParser()
        self.configg = configg
          
    def FILE_WRITE(self, rdic, hld, vld, fnt, fnm, spbtline, gtothead, comp, getfile=None):
        self.configg.add_section('rdic')
        self.configg.add_section('hld')
        self.configg.add_section('vld')
        self.configg.add_section('fonts')
        self.configg.add_section('ftname')
        self.configg.add_section('othinfo')
        self.configg.set('othinfo', 'spbtline', spbtline)
        self.configg.set('othinfo', 'gtothead', gtothead)
        dirpath = rmss_config.my_icon('resources')
        for key in rdic.keys():
            #mval = [rdic[key][0], '']
            self.configg.set('rdic', key, rdic[key])
        for k1 in hld.keys():
            self.configg.set('hld', k1, hld[k1])
        for k2 in vld.keys():
            self.configg.set('vld', k2, vld[k2])
        for k3 in fnt.keys():
            self.configg.set('fonts', k3, fnt[k3])
        for k4 in fnm.keys():
            self.configg.set('ftname', k4, fnm[k4])
        if comp == True:
            if not getfile:
                getfile = self.PGPATH2
            try:
                #with open(rmss_config.my_icon('resources\\pgcompsett.ini'), 'w') as f:
                with open(rmss_config.my_icon(getfile), 'w') as f:
                    self.configg.write(f)
            except IOError as e:
                
                os.makedirs(dirpath)
        else:
            try:
                if not getfile:
                    getfile = self.stfile
                with open(getfile, 'w') as f:
                    self.configg.write(f)
            except IOError as e:
                os.makedirs(dirpath)
    
    def READFILE(self, getfile=None):
        try:
            if not getfile:
                getfile = self.stfile
            configg = ConfigParser.SafeConfigParser()
            configg.read(getfile)
            diclst = self.READCOMMON(configg)
            return diclst
        except :
            return [(),(),]

    def XXX_READFILE(self, getfile=None):
        
            if not getfile:
                getfile = self.stfile
            configg = ConfigParser.SafeConfigParser()
            configg.read(getfile)
            diclst = self.READCOMMON(configg)
            return diclst
        
        
    def READFILE_COMP(self, getfile=None):
        try:
            if not getfile:
                getfile = self.PGPATH2
            configg = ConfigParser.SafeConfigParser()
            #configg.read(rmss_config.my_icon('resources\\pgcompsett.ini'))
            configg.read(rmss_config.my_icon(getfile))
            diclst = self.READCOMMON(configg)
            return diclst
        except :
            return [(),(),]
    def READCOMMON(self, configg):
        vrdic = configg._sections['rdic']
        vhld = configg._sections['hld']
        vvld = configg._sections['vld']
        vfnt = configg._sections['fonts']
        vfname = configg._sections['ftname']
        rdic, hld, vld, fnt, fnm = {}, {}, {}, {}, {}
        for k, v in vrdic.iteritems():
            #rdic[k] = eval(v)
            rdic[k.upper()] = [m for m in eval(v)]
        for k, v in vhld.iteritems():
            hld[k.upper()] = eval(v)
        for k, v in vvld.iteritems():
            vld[k.upper()] = eval(v)
        for k, v in vfnt.iteritems():
            try:
                v = float(v)
            except ValueError:
                v = v
            fnt[k] = v
        for k, v in vfname.iteritems():
            fnm[k] = v
        spbtline = float(configg.get('othinfo','spbtline'))
        gtothead = configg.get('othinfo','gtothead')
        diclst = [rdic, hld, vld, fnt, fnm, spbtline, gtothead]
        return diclst
###########################################################################
#### SMALL CODES ###
def ITMTAXCOLLECT(d, txhdct, itmdct, taxmpkey, txsmhd, igstbool):
        taxr, dted, damt, ttax, txam, ttot, tamt = 0, 0, 0, 0, 0, 0, 0
        dis = 0
        txlst = []
        txdict, txdisc, txaamt = defaultdict(list), defaultdict(list), defaultdict(list)
        tdx = 0
        mtaxr = [28.0, 18.0, 12.0, 5.0,]
        qtycount = []
        for k, v in itmdct.iteritems():
            qtycount.append(float(v['IQTY']))
            try:
                bnv =  float(v['IBON'])
            except ValueError:
                bnv = 0
            qtycount.append(bnv)
            tamt, dis, taxr = float(v[taxmpkey[0]]), v[taxmpkey[1]], float(v[taxmpkey[2]])
            ##tamt, dis = float(v[taxmpkey[0]]), v[taxmpkey[1]]
            dted = float(tamt)*(sum([100.0, -float(dis)])/100.0)
            damt = sum([float(tamt), -dted])
            ttax = dted*(sum([100.0, float(taxr)])/100.0)
            txam = sum([ttax, -dted])
            #print ttax, '>>>', ttax, dted
            txdict[taxr].append(txam)
            txdisc[taxr].append(damt)
            txaamt[taxr].append(tamt)
            tdx += 1
        
        ft1 = dict((k, sum(v)) for k, v in txdict.iteritems())
        ft2 = dict((k, sum(v)) for k, v in txdisc.iteritems())
        ft3 = dict((k, sum(v)) for k, v in txaamt.iteritems())

        fd = {k: (tuple(fd[k] for fd in [ft1, ft2, ft3])) for k in ft1.iterkeys() }
        idx = 0
        rdct = {i.keys()[0] : i.values()[0][0][1] for i in txhdct}
        mvalst = []
        for m, n in fd.iteritems():
            tr = m
            igst = format(float(n[0]), '0.2f')
            ttap = igst
            cgst = format(float(n[0])/2.0, '0.2f')
            sgst = cgst
            damt = format(float(n[1]), '0.2f')
            tamt = format(float(n[2]), '0.2f')
            #ttot = sum([float(n[0]), float(n[2])])
            ttot = sum([float(n[0]), float(n[2]), -float(n[1])])
            ttot = format(float(ttot), '0.2f')
            if igstbool == False:
                mval = {'BTXC':str(tr), 'BIGT':'0.00', 'BCGT':cgst, 'BSGT':sgst, 'TAXP':ttap, 'BDIS':damt, 'BAMT':tamt, 'BTOT':ttot}
            else:
                mval = {'BTXC':str(tr), 'BIGT':igst, 'BCGT':'0.00', 'BSGT':'0.00', 'TAXP':ttap, 'BDIS':damt, 'BAMT':tamt, 'BTOT':ttot}
            mvalst.append(mval)
        sumtax = sum([float(x[0]) for x in fd.values()])
        sumdis = sum([float(x[1]) for x in fd.values()])
        sumamt = sum([float(x[2]) for x in fd.values()])
        if igstbool == False:
            sumigst = '0.00'
            sumcgst = format(sumtax/2.0, '0.2f')
            sumsgst = sumcgst
        else:
            sumigst = format(sumtax, '0.2f')
            sumcgst = '0'
            sumsgst = '0'
        totaxpd = format(sumtax, '0.2f')
        totdisc = format(sumdis, '0.2f')
        totamt = format(sumamt, '0.2f')
        gttotl = format(sum([sumtax, sumamt, -sumdis]), '0.2f')
        gttot_adj = sum([sumtax, sumamt, -sumdis])
        gttotlall = CountOtherVales(d['OTH1'][1], d['OTH2'][1], d['OTH3'][1], gttot_adj)
        
        gdtotlr = format(round(round(float(gttotlall),2)), '0.2f')
        gtotlst = [sumigst, sumcgst, sumsgst, totaxpd, totdisc, totamt, gttotl, gttot_adj, gttotlall, gdtotlr]
    
        return fd, rdct, qtycount, mvalst, gtotlst
    
def CountOtherVales(oth1, oth2, oth3, tot):
        othr1, othr2, othr3 = OthSplit(oth1), OthSplit(oth2), OthSplit(oth3)
        gtot = sum([ChkVal(othr1), ChkVal(othr2), ChkVal(othr3), float(tot)])
        return gtot
def OthSplit(oth):
        try:
            othd = oth.split(':')[1]
        except IndexError:
            othd = 0
        
        return othd
def ChkVal(oth):
        try:
            oth = float(oth)
        except ValueError:
            oth = 0
        return oth
def OwnrSet(dcf, odet):
        owner, o_add1, o_add2, phone, o_dl_no, o_tin, statu1, statu2, statu3, statu4, oemail = odet
        dcf['ONDP'][0][0] = owner
        dcf['ONDP'][0][0] = owner
        dcf['OAD1'][0][0] = o_add1.title()
        dcf['OAD2'][0][0] = o_add2.title()
        dcf['ONDP'][1] = ''
        dcf['OAD1'][1] = ''
        dcf['OAD2'][1] = ''
        dcf['OPHN'][1] = phone
        dcf['OREG'][1] = o_dl_no
        dcf['OGST'][1] = o_tin
        dcf['OEML'][1] = oemail[0]
        dcf['STA1'][1] = ' '.join([statu1.capitalize(),statu2.capitalize()])
        dcf['STA2'][1] = ' '.join([statu3.capitalize(),statu4.capitalize()])
        dcf['STA3'][1] = ''
        dcf['STA4'][1] = ''
        
##################################
blst = []
def odt(bn):
    "Give Bill Number"
    if blst == []:
        mess = 'Original'
        blst.append(bn)
    else:
        if blst[0] == bn:
            mess = 'Duplicate'
            blst.append(bn)
        if blst[0] != bn:
            mess = 'Original'
            del blst[:]
            blst.append(bn)
            return mess
        try:
            if blst[2] == bn:
                mess = 'Triplicate'
                del blst[:]   
        except:
            mess = 'Duplicate'
    return mess

##################################
class MRECEIPT_PDF(object):
    def __init__ (self, compval, resource_dic, pdf_name, itm, BMAIN, dcf, dhl, dvl, fnt, fnm, spbtline, gtstx, igstbool, pdf_open, parent=None):
        self.resource_dic = resource_dic
        #if len(itm) > 30:
        #    self.MLPGPrint(compval, resource_dic, pdf_name, itm, BMAIN, dcf, dhl, dvl, fnt, fnm, spbtline, gtstx, igstbool, pdf_open)
        #else:
        self.RegularPrint(compval, resource_dic, pdf_name, itm, BMAIN, dcf, dhl, dvl, fnt, fnm, spbtline, gtstx, igstbool, pdf_open, parent=parent)
            
    def MLPGPrint(self, compval, resource_dic, pdf_name, itm, BMAIN, dcf, dhl, dvl, fnt, fnm, spbtline, gtstx, igstbool, pdf_open, parent=None):
        self.compval = compval
        self.igstbool = igstbool
        self.dcf, self.dhl, self.dvl = dcf,  dhl, dvl
        self.BMAIN = BMAIN
        PG_RIGHT_LINE_POS = 10
        self.HL_BUFF_WDTH = 1.063
        mfpath = os.path.abspath(rmss_config.PP_EXPORT)#, rmss_config.PP_EXPORT,'rmss_config.PP_EXPORT'
        self.export_path ="%s\\%s"%(mfpath, pdf_name)
        try:
            owner,o_add1,o_add2,phone,o_dl_no,o_tin,statu1,statu2,statu3,statu4, osn, osoth = resource_dic['owner']
        except :
            owner, o_add1, o_add2, phone = 'RMS DEMO ', 'RMS ADDRESS 1', 'RMS ADDRESS 2','9999999999'
            o_dl_no = 'DEOS-2017/20/00093'     
            o_tin = '09AVDPM1996P1ZP'
            statu1,statu2,statu3,statu4 = 'statu1','statu2','statu3','statu4'
        stklst = '1,2,3..... Stockist List.....'
        ###########################################################
        #fnt = dcf.get('FONTSIZE')
        fs1 = fnt.get('fs1',7)
        ftitm = fnt.get('fit',8)
        fitmh = fnt.get('fitmh',9)
        f2pt = fnt.get('f2pt',9)
        fh4 = fnt.get('fh4',8)
        fh5 = fnt.get('fh5',7)
        fh6 = fnt.get('fh6',8)
        fh3 = fnt.get('fh3',9)
        fh2 = fnt.get('fh2', 12)
        fh1 = fnt.get('fh1',15)
        
        fh1n = fnm.get('fh1n') 
        fh2n = fnm.get('fh2n') 
        fitn = fnm.get('fitn') 
        fs1n = fnm.get('fs1n') 
        fh3n = fnm.get('fh3n') 
        fitnh = fnm.get('fitnh')
        f2ptn = fnm.get('f2ptn') 
        fh4n = fnm.get('fh4n')
        fh5n = fnm.get('fh5n')
        fh6n = fnm.get('fh6n')
        
        try:
            self.lprt = dhl['LPRT'][0]
            lndmlwidth = dhl['LPRT'][2]
        except:
            self.lprt = 2
            lndmlwidth = 810
        ### Horizontal Line Top Side ###################################
        pglftht = dcf.get('PAGESIZE')
        lnpos = sum([pglftht[1], -13])
        lwidth = sum([pglftht[0], -15])
        mlwidth = sum([lwidth,-PG_RIGHT_LINE_POS]) ## Modified Horizontal Line Width
        #print dcf.get('PAGESIZE')
        #print dcf.get('ONDP')
        if self.lprt == 2:
            cv = canvas.Canvas(self.export_path, pagesize=portrait(dcf.get('PAGESIZE')))
        else:
            try:
                mlwidth = int(lndmlwidth)
            except:
                mlwidth = 0
            cv = canvas.Canvas(self.export_path,pagesize=landscape(dcf.get('PAGESIZE')))
        cv.setLineWidth(0.3)
    
        self.LOGOSET(cv, dcf, fh1n, fh1,'LOGO')
        try:
            oemail = resource_dic['oemail']
            odet = owner, o_add1, o_add2, phone, o_dl_no, o_tin, statu1, statu2, statu3, statu4, oemail
            OwnrSet(dcf, odet)
        except:
            return
        dcf = self.BMainChk(dcf)
        
        odtcscr_o = dcf['THDP'][1]
        
        dcf['THDP'][1] = ''.join([odtcscr_o, '   [',odt(dcf['BILN'][1]), '-Copy]'])
            
        self.VCV(cv, dcf, 'Times-Roman', 6, 'THDP')
        
        self.BOXVCV(cv, dcf, fh4n, fh4,'CSCR')
        #### Font Name, Size is Same For Owner Address And Party Address, to Avaoid Confusion 02/06/2020, f2ptn, f2pt
        self.VCV(cv, dcf, fh1n, fh1,'ONDP')
        ownlst = ['OAD1', 'OAD2', 'OPHN', 'OEML', 'OREG', 'OGST',]
        
        for o in range(len(ownlst)):
            ###self.VCV(cv, dcf, fh2n, fh4,ownlst[o])
            self.VCV(cv, dcf, fh2n, fh4, ownlst[o])
            
        self.VCV(cv, dcf, fh2n, fh2,'PNDP')
        self.VCV(cv, dcf, fs1n, fh3,'BILN')
        self.VCV(cv, dcf, fs1n, fh3,'BDAT')
        
        prtlst = ['PAD1', 'PAD2', 'PMOD', 'BINF', 'PPHN', 'PREG', 'PGST', 'DITM', 'DPNO', 'PEML', 'ORDN', 'DPTO']
        for p in range(len(prtlst)):
            self.VCV(cv, dcf, f2ptn, f2pt, prtlst[p])  
        
        sitmlst = ['HSNC', 'ITMN', 'IPAC', 'IQTY', 'IBON', 'IRAT',
                   'IRA2', 'IAMT', 'ITAX', 'IDIS', 'IMRP', 'IEXP', 'INET', 'IBAT']
        taxmpkey = ['IAMT', 'IDIS', 'ITAX', 'INET'] 
        icol = []
        for il in sitmlst:
            try:
                icol.append(dcf.get(il)[0][1][0])
            except IndexError:
                pass
        try:
            line = sum([dcf.get('ITMN')[0][1][1], -18.0])
        except :
            return
        
        itmdct = {}
        for n, m in enumerate(itm):
            itmdct[n] = {}
            for i in range(len(sitmlst)):
                itmdct[n][sitmlst[i]] = itm[n][i]
        
        itmhdct = self.HSCV(cv, dcf, fitnh, fitmh, sitmlst)
        
        ########################################################
        self.pght = dcf.get('PAGESIZE')[1]
        self.pgwt = dcf.get('PAGESIZE')[0]
        ########################################################
        
        ### cv.showPage() ==>> will create new page
        self.ITMDCV_MLPG(cv, dcf, fitn, ftitm, itm, line, icol, sitmlst, itmdct, itmhdct, spbtline)
        ########################################################
        ### Automatic Calculation And Adjustments of Item Dispaly
        ### if Item Numbers are greater than given parameters
        itemcount = len(itm)
        itmstart = sum([dcf['ITMN'][0][1][1], -10])
        itmend = sum([dcf['BTXC'][0][1][1], 10])
        
        itmendpos = self.ItemAreaCal(itemcount, spbtline, itmstart, itmend, ftitm, )
        try:
            if itmendpos > 0:
                mitmend = sum([itmend,20])
                self.mdcf = copy.deepcopy(dcf)
                self.mdvl = copy.deepcopy(dvl)
                self.mdhl = copy.deepcopy(dhl)
                rdickeylst = self.GetResetRdicKeys(mitmend, dcf)
                vlkeylst = self.GetResetVLKeys(mitmend, dvl)
                hlkeylst = self.GetResetHLKeys(mitmend, dhl)
                self.ResetDicPos(itmendpos, rdickeylst, vlkeylst, hlkeylst)
            
                dcf = self.mdcf
                dvl = self.mdvl
                dhl = self.mdhl
        except :  
            pass
       
        ########################################################
        if self.compval == False:
            txsmhd = ['BTXC','BIGT','BCGT','BSGT', 'TAXP', 'BDIS','BAMT', 'BTOT']
        else:
            dcf['BTXC'][0][0] = ""
            dcf['BIGT'][0][0] = ""
            dcf['BCGT'][0][0] = ""
            dcf['BSGT'][0][0] = ""
            dcf['TAXP'][0][0] = ""
            dcf['BDIS'][0][0] = ""
            dcf['BAMT'][0][0] = ""
            dcf['BTOT'][0][0] = ""
            txsmhd = ['BTXC','BIGT','BCGT','BSGT', 'TAXP', 'BDIS','BAMT', 'BTOT']
        
        txhdct = self.HSCV(cv, dcf, fh2n, ftitm, txsmhd)
       
        btlst = self.BTXCV_MLPG(cv, dcf, fh6n, fh6, txhdct, itmdct, taxmpkey, txsmhd, 10, 215)### spbtline is constant/fixed here == 10
        ## Line Start Space, Line Start Col, Line Width, Line End Col
        ################################################################
        
        dcf['ITMC'][1] = str(btlst[0])
        dcf['QTTC'][1] = str(btlst[1])
        dcf['DISS'][1] = str(btlst[2])
        dcf['AMTS'][1] = str(btlst[3])
        if self.compval == True:
            dcf['AMTS'][1] = format(sum([float(btlst[4]), float(btlst[2])]), '0.2f')

        dcf['GTOT'][1] = ' '.join([gtstx,str(btlst[4])])
        #dcf['GTOT'][1] = str(btlst[4])
        amtwrd = num2word.to_card(round(float(btlst[4])))
        dcf['AWRD'][1] = amtwrd
        
        ###############################################
        mawrdfsize = fh4
        if len(amtwrd)> 50:
            mawrdfsize  = fh4-1
        
        adjrund = format("%.2f" % float(btlst[5]))
        
        dcf['OTH3'][1] = adjrund
        btmlst = ['DISS', 'AMTS', 'OTH1', 'OTH2', 'OTH3']
        lr, lc = 400, 215 ## Verical Flow
        self.VCV_MLPG(cv, dcf, fh4n, fh4,'ITMC', lr, lc)
        self.VCV_MLPG(cv, dcf, fh4n, fh4,'QTTC', 450, lc)
        lc = lc-15
        for b in range(len(btmlst)):
            self.VCV_MLPG(cv, dcf, fh3n, fh3, btmlst[b], lr, lc)
            lc -= 15
        try:
            gtfont = sum([fh3,2])
        except :
            gtfont = fh3
        self.BOXVCV_MLPG(cv, dcf, fs1n, gtfont,'GTOT', lr, lc)
        lc = lc-10 
        self.VCV_MLPG(cv, dcf, fs1n, mawrdfsize, 'AWRD', 20, lc)
        lc = lc-10 
        self.VCV_MLPG(cv, dcf, fh3n, fh4, 'CMNT', 20, lc)
        btmlst = ['STA1', 'STA2', 'STA3', 'STA4']
        for b in range(len(btmlst)):
            lc -= 10
            self.VCV_MLPG(cv, dcf, fh5n, fh5, btmlst[b], 20, lc)
        lc -= 10
        self.VCV_MLPG(cv, dcf, fh4n, fh4, 'AUSF', 290, lc)
        dcf['AUSO'][0][0] = owner
        dcf['AUSO'][1] = ''
        self.VCV_MLPG(cv, dcf, fitnh, fitmh, 'AUSO', 310, lc)
        self.VCV_MLPG(cv, dcf, fitnh, fitmh, 'AUSS', 325, lc)
        cv.setFont('Courier',7)
        sunilb = dvl['_END0'][1]
        sunilh = sum([(dcf.get('PAGESIZE')[0]/2.0),-50])
        cv.drawString(20, 30,'RMS SOFT : www.rmssoft.co.in')
        cv.save()
        
        if pdf_open == True:
            try:
                exppath = 'start %s'%self.export_path
                dc = subprocess.Popen(exppath,shell=True)
                dc.wait()
            except :
                pass
    def RegularPrint(self, compval, resource_dic, pdf_name, itm, BMAIN, dcf, dhl, dvl, fnt, fnm, spbtline, gtstx, igstbool, pdf_open, parent=None):
        
        self.compval = compval
        self.igstbool = igstbool
        self.dcf, self.dhl, self.dvl = dcf,  dhl, dvl
        self.BMAIN = BMAIN
       
        PG_RIGHT_LINE_POS = 10
        self.HL_BUFF_WDTH = 1.063
        mfpath = os.path.abspath(rmss_config.PP_EXPORT)#, rmss_config.PP_EXPORT,'rmss_config.PP_EXPORT'
        self.export_path ="%s\\%s"%(mfpath, pdf_name)
        try:
            owner,o_add1,o_add2,phone,o_dl_no,o_tin,statu1,statu2,statu3,statu4, osn, osoth = resource_dic['owner']
        except :
            owner, o_add1, o_add2, phone = 'RMS DEMO PARTY NAME', 'RMS DEMO ADDRESS1', 'RMS DEMO ADDRESS2','9999999999'
            o_dl_no = 'DEOS-2021/20/12393'     
            o_tin = '99XYZPM1234P1ZP'
            statu1,statu2,statu3,statu4 = 'statu1','statu2','statu3','statu4'
        stklst = '1,2,3..... Stockist List.....'
        ###########################################################
        #fnt = dcf.get('FONTSIZE')
        fs1 = fnt.get('fs1',7)
        ftitm = fnt.get('fit',8)
        fitmh = fnt.get('fitmh',9)
        f2pt = fnt.get('f2pt',9)
        fh4 = fnt.get('fh4',8)
        fh5 = fnt.get('fh5',7)
        fh6 = fnt.get('fh6',8)
        fh3 = fnt.get('fh3',9)
        fh2 = fnt.get('fh2', 12)
        fh1 = fnt.get('fh1',15)
        
        fh1n = fnm.get('fh1n') 
        fh2n = fnm.get('fh2n') 
        fitn = fnm.get('fitn') 
        fs1n = fnm.get('fs1n') 
        fh3n = fnm.get('fh3n') 
        fitnh = fnm.get('fitnh')
        f2ptn = fnm.get('f2ptn') 
        fh4n = fnm.get('fh4n')
        fh5n = fnm.get('fh5n')
        fh6n = fnm.get('fh6n')
        
        try:
            self.lprt = dhl['LPRT'][0]
            lndmlwidth = dhl['LPRT'][2]
        except:
            self.lprt = 2
            lndmlwidth = 810
        ### Horizontal Line Top Side ###################################
        pglftht = dcf.get('PAGESIZE')
        try:
            ### Page Top Side Horizontal Line Position
            toplinemarg = (pglftht[1]-dcf['THDP'][0][1][1])-7
            lnpos = sum([pglftht[1], -toplinemarg])
        except:
            lnpos = sum([pglftht[1], -13])
        lwidth = sum([pglftht[0], -15])
        mlwidth = sum([lwidth,-PG_RIGHT_LINE_POS]) ## Modified Horizontal Line Width
        
        if self.lprt == 2:
            try:
                gtfont = sum([fh3,2]) ### Extra Font Size Used in Grand Totla Specially For Potrait 
            except :
                gtfont = fh3
            cv = canvas.Canvas(self.export_path, pagesize=portrait(dcf.get('PAGESIZE')))
        else:
            try:
                mlwidth = int(lndmlwidth)
            except:
                mlwidth = 0
            try:
                gtfont = sum([fh3,5]) ### Extra Font Size Used in Grand Totla Specially For Landscape 
            except :
                gtfont = fh3
            cv = canvas.Canvas(self.export_path,pagesize=landscape(dcf.get('PAGESIZE')))
            
        cv.setLineWidth(0.3)
    
        self.LOGOSET(cv, dcf, fh1n, fh1,'LOGO')
        try:
            oemail = resource_dic['oemail']
            odet = owner, o_add1, o_add2, phone, o_dl_no, o_tin, statu1, statu2, statu3, statu4, oemail
            OwnrSet(dcf, odet)
        #except Exception as err:
        except KeyboardInterrupt as err:
            wx.MessageBox('Some Setting has been changed or corrupt by User\n'
                          'in PRINT BILL PAGE SETTING, this create trouble\n'
                          'while printing or Exporting SALE_BILL.pdf\n\n'
                          'Go to Master >> PRINT BILL PAGE SETTING and \nCreate New Settings.',
                          'RMS PRINT BILL PAGE SETTING ERROR',wx.ICON_EXCLAMATION)
            if parent:
                parent.parent.status.SetStatusText(" [pagesetup, line 559 ERROR]: You have Changed or Deleted Some "
                "TAGS in PRINT BILL PAGE SETTING, Generate ERROR While Print/Export Bills")
            return
        dcf = self.BMainChk(dcf)
        
        odtcscr_o = dcf['THDP'][1]
        
        dcf['THDP'][1] = ''.join([odtcscr_o, '   [',odt(dcf['BILN'][1]), '-Copy]'])
            
        self.VCV(cv, dcf, 'Times-Roman', 6, 'THDP')
        #### Font Name, Size is Same For Owner Address And Party Address, CSCR (relatively incremented by 1)
        #### to Avaoid Confusion 02/06/2020, f2ptn, f2pt
        self.BOXVCV(cv, dcf, fh4n, sum([f2pt,1]),'CSCR')
        
        self.VCV(cv, dcf, fh1n, fh1,'ONDP')
        ownlst = ['OAD1', 'OAD2', 'OPHN', 'OEML', 'OREG', 'OGST',]
        
        for o in range(len(ownlst)):
            ###self.VCV(cv, dcf, fh2n, fh4,ownlst[o])
            self.VCV(cv, dcf, f2ptn, f2pt, ownlst[o])
            
        self.VCV(cv, dcf, fh2n, fh2,'PNDP')
        self.VCV(cv, dcf, fs1n, fh3,'BILN')
        self.VCV(cv, dcf, fs1n, fh3,'BDAT')
        
        prtlst = ['PAD1', 'PAD2', 'PMOD', 'BINF', 'PPHN', 'PREG', 'PGST', 'DITM', 'DPNO', 'PEML', 'ORDN', 'DPTO']
        
        for p in range(len(prtlst)):
            self.VCV(cv, dcf, f2ptn, f2pt, prtlst[p])  
        
        sitmlst = ['HSNC', 'ITMN', 'IPAC', 'IQTY', 'IBON', 'IRAT',
                   'IRA2', 'IAMT', 'ITAX', 'IDIS', 'IMRP', 'IEXP', 'INET', 'IBAT']
        taxmpkey = ['IAMT', 'IDIS', 'ITAX', 'INET'] 
        icol = []
        for il in sitmlst:
            if isinstance (dcf.get(il)[0], list):
                ### if any column removed by user, will be excluded in this iteration 
                try:
                    icol.append(dcf.get(il)[0][1][0])
                except IndexError:
                    pass
        try:
            line = sum([dcf.get('ITMN')[0][1][1], -18.0])
        except :
            return
        
        itmdct = {}
        for n, m in enumerate(itm):
            itmdct[n] = {}
            for i in range(len(sitmlst)):
                itmdct[n][sitmlst[i]] = itm[n][i]
        
        itmhdct = self.HSCV(cv, dcf, fitnh, fitmh, sitmlst)
        ########################################################
        self.pght = dcf.get('PAGESIZE')[1]
        self.pgwt = dcf.get('PAGESIZE')[0]
        ########################################################
        if self.resource_dic['iteminfo']:
            self.ITMDCV_iteminfo(cv, dcf, fitn, ftitm, itm, line, icol, sitmlst, itmdct, itmhdct, spbtline)
        else:
            self.ITMDCV_default(cv, dcf, fitn, ftitm, itm, line, icol, sitmlst, itmdct, itmhdct, spbtline)
        ### Automatic Calculation And Adjustments of Item Dispaly
        ### if Item Numbers are greater than given parameters
        itemcount = len(itm)
        itmstart = sum([dcf['ITMN'][0][1][1], -10])
        itmend = sum([dcf['BTXC'][0][1][1], 10])
        
        itmendpos = self.ItemAreaCal(itemcount, spbtline, itmstart, itmend, ftitm, )
        try:
            if itmendpos > 0:
                mitmend = sum([itmend,20])
                self.mdcf = copy.deepcopy(dcf)
                self.mdvl = copy.deepcopy(dvl)
                self.mdhl = copy.deepcopy(dhl)
                rdickeylst = self.GetResetRdicKeys(mitmend, dcf)
                vlkeylst = self.GetResetVLKeys(mitmend, dvl)
                hlkeylst = self.GetResetHLKeys(mitmend, dhl)
                self.ResetDicPos(itmendpos, rdickeylst, vlkeylst, hlkeylst)
            
                dcf = self.mdcf
                dvl = self.mdvl
                dhl = self.mdhl
        except :  
            pass
       
        ########################################################
        if self.compval == False:
            txsmhd = ['BTXC','BIGT','BCGT','BSGT', 'TAXP', 'BDIS','BAMT', 'BTOT']
        else:
            dcf['BTXC'][0][0] = ""
            dcf['BIGT'][0][0] = ""
            dcf['BCGT'][0][0] = ""
            dcf['BSGT'][0][0] = ""
            dcf['TAXP'][0][0] = ""
            dcf['BDIS'][0][0] = ""
            dcf['BAMT'][0][0] = ""
            dcf['BTOT'][0][0] = ""
            txsmhd = ['BTXC','BIGT','BCGT','BSGT', 'TAXP', 'BDIS','BAMT', 'BTOT']
        
        txhdct = self.HSCV(cv, dcf, fh2n, ftitm, txsmhd)
        btlst = self.BTXCV(cv, dcf, fh6n, fh6, txhdct, itmdct, taxmpkey, txsmhd, 10)### spbtline is constant/fixed here == 10
        ## Line Start Space, Line Start Col, Line Width, Line End Col
        ################################################################
        
        dcf['ITMC'][1] = str(btlst[0])
        dcf['QTTC'][1] = str(btlst[1])
        dcf['DISS'][1] = str(btlst[2])
        dcf['AMTS'][1] = str(btlst[3])
        if self.compval == True:
            dcf['AMTS'][1] = format(sum([float(btlst[4]), float(btlst[2])]), '0.2f')

        dcf['GTOT'][1] = ' '.join([gtstx,str(btlst[4])])
        #dcf['GTOT'][1] = str(btlst[4])
        amtwrd = num2word.to_card(round(float(btlst[4])))
        dcf['AWRD'][1] = amtwrd
        
        ###############################################
        mawrdfsize = fh4
        if len(amtwrd)> 50:
            mawrdfsize  = fh4-1
        self.VCV(cv, dcf, fh4n, fh4,'ITMC')
        self.VCV(cv, dcf, fh4n, fh4,'QTTC')
        adjrund = btlst[5]
        
        dcf['OTH3'][1] = format("%.2f" % float(btlst[5]))
        btmlst = ['DISS', 'AMTS', 'OTH1', 'OTH2', 'OTH3']
        for b in range(len(btmlst)):
            self.VCV(cv, dcf, fh3n, fh3, btmlst[b])   
        
        self.BOXVCV(cv, dcf, fs1n, gtfont,'GTOT')
        
        self.VCV(cv, dcf, fs1n, mawrdfsize, 'AWRD')
        self.VCV(cv, dcf, fh3n, fh4, 'CMNT')
        
        btmlst = ['STA1', 'STA2', 'STA3', 'STA4']
        for b in range(len(btmlst)):
            self.VCV(cv, dcf, fh5n, fh5, btmlst[b])
        self.VCV(cv, dcf, fh4n, fh4, 'AUSF')
        #dcf['AUSO'][0][0] = owner
        dcf['AUSO'][1] = ''
        
        self.VCV(cv, dcf, fitnh, fitmh, 'AUSO')
        self.VCV(cv, dcf, fitnh, fitmh, 'AUSS')
        ### Horizontal Tax Box, Bottom Box Lines
        self.AHL(cv, dhl)
        try:
            ### lincol = Starting Left PageSide Position
            ### endln = Vertical Page End Position
            ### endstr = Empty String not using yet
            lincol, endln, endstr = dvl['_END0']
        except:
            lincol, endln, endstr = 10, 10, ""
        ### Horizontal Top Line PAGE BOX Boundary
        cv.line(lincol, lnpos, mlwidth, lnpos)
        ### Horizontal Line Bottom Side
        
        
        ### Vertical Line Left Side PAGE BOX Boundary
        cv.line(lincol, endln, lincol, lnpos)
        ### Vertical Line Right Side PAGE BOX Boundary
        cv.line(mlwidth , endln, mlwidth, lnpos)
        ### Horizontal Bottom Line PAGE BOX Boundary 
        cv.line(lincol , endln, mlwidth, endln)
        ###cv.setDash(6,3)
        ###cv.setStrokeGray(0.8)
        cv.setStrokeGray(0.5) ### Setting gray / Gray color vertical lines inside Page Border
        self.AVL(cv, dvl)  ### VERTICAL LINE for ITEM DISPLAY
    
        cv.setFont('Courier',7)
        sunilb = sum([dvl['_END0'][1],2])
        sunilh = sum([(dcf.get('PAGESIZE')[0]/2.0),-50])
        cv.drawString(sunilh, sunilb,'RMS SOFT :www.rmssoft.co.in')
        cv.setFont('Courier',13)
        cv.drawString(sunilh-10, sunilb-5,'\xc2\xa9')
        ### for raw string '\u20B9' INR
        cv.save()
        
        if pdf_open == True:
            try:
                ##reqfpath = ffalldrive(drive='C:\\', prymfolder='Program Files',
                ##            searchfolder='Adobe', filename='AcroRd32.exe')
                ##printername = win32print.GetDefaultPrinter()
                exppath = 'start %s'%self.export_path
                dc = subprocess.Popen(exppath,shell=True)
                dc.wait()
            except  :
                pass
    def BMainChk(self, dcf):
        if self.BMAIN == True:
            bmdcf = dcf
            bmdcf['AUSO'][0][0] = dcf['ONDP'][0][0]
        else:
            compdic = {False:"Estimate/Quotation", True:"Estimate/Quotation"}
            bmdcf = copy.deepcopy(dcf)
            if self.resource_dic['onm_on_esti_bill'].upper()=='TRUE':
                bmdcf['AUSO'][0][0] = dcf['ONDP'][0][0]
            elif self.resource_dic['onm_on_esti_bill'].upper()=='FALSE':
                bmdcf['AUSO'][0][0]=''
                bmdcf['ONDP'][0] = ''
                bmdcf['OAD1'][0] = ''
                bmdcf['OAD2'][0] = ''
                bmdcf['OPHN'][0] = ''
                bmdcf['OEML'][0] = ''
                bmdcf['OREG'][0] = ''
                bmdcf['OGST'][0] = ''
                bmdcf['OPHN'][0] = ''
                bmdcf['AUSS'][0] = ''
                bmdcf['AUTS'][0] = ''
                bmdcf['BILN'][0][0] = ''
                bmdcf['BDAT'][0][0] = ''
            elif self.resource_dic['onm_on_esti_bill'].upper()=='PART1':
                bmdcf['OREG'][0] = ''
                bmdcf['OGST'][0] = ''
                bmdcf['AUSO'][0][0] = dcf['ONDP'][0][0]
            elif self.resource_dic['onm_on_esti_bill'].upper()=='PART2':
                bmdcf['AUSO'][0][0] = dcf['ONDP'][0][0]
                bmdcf['OAD1'][0] = ''
                bmdcf['OAD2'][0] = ''
                bmdcf['OPHN'][0] = ''
                bmdcf['OEML'][0] = ''
                bmdcf['OREG'][0] = ''
                bmdcf['OGST'][0] = ''
                bmdcf['OPHN'][0] = ''
                bmdcf['BILN'][0][0] = ''
                bmdcf['BDAT'][0][0] = ''
            else:
                bmdcf['AUSO'][0][0]=''
                bmdcf['ONDP'][0] = ''
                bmdcf['OAD1'][0] = ''
                bmdcf['OAD2'][0] = ''
                bmdcf['OPHN'][0] = ''
                bmdcf['OEML'][0] = ''
                bmdcf['OREG'][0] = ''
                bmdcf['OGST'][0] = ''
                bmdcf['OPHN'][0] = ''
                bmdcf['AUSS'][0] = ''
                bmdcf['AUTS'][0] = ''
                bmdcf['BILN'][0][0] = ''
                bmdcf['BDAT'][0][0] = ''
        return bmdcf
    def GetResetRdicKeys(self, mitmend, dcf):
        rdickeylst = []
        for x, y in dcf.iteritems():
            try:
                if mitmend >= float(y[0][1][1]):
                    rdickeylst.append(x)
            except :
                pass
        return rdickeylst
    def GetResetVLKeys(self, mitmend, dvl):
        vlkeylst = []
        for x, y in dvl.iteritems():
            try:
                if mitmend >= float(y[1]):
                    vlkeylst.append(x)
            except (ValueError, KeyError):
                pass
        return vlkeylst
    
    def GetResetHLKeys(self, mitmend, dhl):
        hlkeylst = []
        for p, q in dhl.iteritems():
            try:
                if mitmend >= float(q[1]):
                    hlkeylst.append(p)
            except (ValueError, KeyError):
                pass
        return hlkeylst
    def ItemAreaCal(self, itemcount, spbtline, itmstart, itmend, ftitm, ):
        #item End is the position of Bottom GST Tax Heading is given
        #Calculate Area Occupied By Given Items """

        itmsize = itemcount*1.20
        itmgap = sum([itemcount*spbtline])
        itmarea = sum([itmgap, itmsize, -(spbtline*2)])
        itmendpos = sum([itmarea, -sum([itmstart, -itmend])])
        
        return itmendpos
    
    def ResetDicitmPos(self, ds, prvpos, itmendpos):
        repos = sum([prvpos, -itmendpos])
        self.mdcf[ds][0][1][1] = repos 
        return self.mdcf
    def ResetVLDic(self, ds, prvpos, itmendpos):
        repos = sum([prvpos, -itmendpos])
        self.mdvl[ds][1] = repos 
        return self.mdvl
    def ResetHLDic(self, ds, prvpos, itmendpos):
        repos = sum([prvpos, -itmendpos])
        self.mdhl[ds][1] = repos 
        return self.mdhl
    
    def ResetDicPos(self, itmendpos, rdickeylst, vlkeylst, hlkeylst):
        
        self.RDic_pos_Reset(rdickeylst, itmendpos)
        self.VLDic_pos_Reset(vlkeylst, itmendpos)
        self.HLDic_pos_Reset(hlkeylst, itmendpos)
        
    def RDic_pos_Reset(self, rdickeylst, itmendpos):
        for ds in rdickeylst:
            try:
                prvpos = self.dcf[ds][0][1][1]
                self.ResetDicitmPos(ds, prvpos, itmendpos)
                #print dcf[ds][0][1][1]
            except IndexError:
                pass
        
    def VLDic_pos_Reset(self, vlkeylst, itmendpos):
        for kv in vlkeylst:
            try:
                prvpos = self.dvl[kv][1]
                self.ResetVLDic(kv, prvpos, itmendpos)
            except:
                pass
        
    def HLDic_pos_Reset(self, hlkeylst, itmendpos):
        for kv in hlkeylst:
            try:
                prvpos = self.dhl[kv][1]
                self.ResetHLDic(kv, prvpos, itmendpos)
            except:
                pass
        
    def LOGOSET(self, cv, d, fn, fz, ds):
        try:
            #cv.drawString(d.get(ds)[0][1][0],d.get(ds)[0][1][1], txd)
            #print d.get(ds)
            w, h = self.ValSplit( d.get(ds)[0][0], )
            cv.drawImage(rmss_config.my_icon('resources\\mylogo.jpg'),d.get(ds)[0][1][0], d.get(ds)[0][1][1]-2, w, h)
            #cv.drawImage(rmss_config.my_icon('resources\\mylogo.jpg'),d.get(ds)[0][1][0], d.get(ds)[0][1][1]-2, width=25, height=25)
        except:
            pass
    def SetLine(self, dl, idx):
        for k, v in sorted(dl.iteritems()):
            try:
                float(v[0])
                dl[k[4:6]] = [dl.get(k.replace('E','S'))[0], dl.get(k.replace('S','E'))[idx],  dl.get(k)[1]]
            except :
                pass

    def AHL(self, cv, dl):
        self.SetLine(dl, 0)
        for k, v in sorted(dl.iteritems()):
            try:
                float(k)
                linepos = sum([v[2], -2.5])
                ### 1.08 is used as line buffer
                linwdth = sum([(v[1]*self.HL_BUFF_WDTH),-v[0]])
                cv.line(v[0] , linepos, linwdth, linepos)
                ## Line Start Space, Line Start Col, Line Width, Line End Col
            except (ValueError, TypeError):
                pass

    def VCV(self, cv, d, fn, fz, ds):
        try:
            cv.setFont(fn,fz)
            txd = ' '.join([d.get(ds)[0][0], d.get(ds)[1]])
            cv.drawString(d.get(ds)[0][1][0],d.get(ds)[0][1][1], txd)
        except:
            pass
    def ValSplit(self, val, ):
        try:
            wh = val.split('*')
            w = float(wh[0])*10
            h = float(wh[1])*10
        except IndexError:
            w, h = 0, 0
        return w, h
    
    def BOXVCV(self, cv, d, fn, fz, ds):
        try:
            wh = d.get(ds)[0][0].split('*')
            w = float(wh[0])*10
            h = float(wh[1])*10
            w, h = self.ValSplit( d.get(ds)[0][0], )
            reccol = d.get(ds)[0][1][0]-5.0
            rcrow = d.get(ds)[0][1][1]-2.5
            #############################
            cv.setFillColorRGB(0.95,0.95,0.95)
            ##cv.rect(rectCOLpos,rectROWpos,'rect_width',rect_height, fill=0)##stroke=0 >> 0=NoBorder, 1=Border fill=0 >>0=None,1=Yes
            cv.rect(reccol,rcrow,w,h, stroke=0, fill=1)
            cv.setFillColorRGB(0,0,0)
            #############################
        except (ValueError, IndexError):
            pass
        try:
            cv.setFont(fn,fz)
            #txd = ' '.join([d.get(ds)[0][0], d.get(ds)[1]])
            txd = d.get(ds)[1]
            cv.drawString(d.get(ds)[0][1][0],d.get(ds)[0][1][1], txd)
        except:
            pass

    def HSCV(self, cv, d, fn, fz, hdlst):
        itmlst = []
        for ds in hdlst:
            try:
                cv.setFont(fn,fz)
                if isinstance(d.get(ds)[0], list):
                    itmlst.append({ds:d.get(ds)})
                    txd = ' '.join([d.get(ds)[0][0], d.get(ds)[1]])
                    cv.drawString(d.get(ds)[0][1][0],d.get(ds)[0][1][1], txd)
            except :
                pass
        return itmlst

    def AVL(self, cv, dl):
        self.SetLine(dl, 1)
        for k, v in sorted(dl.iteritems()):
            try:
                float(k)
                ### -5 is buffer added to shift line col on left side of page
                lhtcol = sum([float(v[0]), -5])
                ### 7 is used as line buffer
                lhtsrt = sum([float(v[2]), 7])
                lhtend = float(v[1])
                cv.line(lhtcol , lhtend, lhtcol, lhtsrt)
            except (ValueError, TypeError):
                pass
            
    def ITMDCV_MLPG(self, cv, d, fn, fz2, itm, line, icol, sitmlst, itmdct, itmhdct, spbtline):
        rdct = {i.keys()[0] : i.values()[0][0][1] for i in itmhdct}
        ihk = itmdct.values()[0].keys()
        idx = 0
        ########################################################################
        lastline = line
        
        usedspc = sum([self.pght, 15, -line]) ### Page Start After Taking 15 point space buffer 
        btmspc = 110.0  ## Pre Defined Value
        itemspc = sum([usedspc, btmspc, 30])
        ##pagemaxline = sum([(int(itemspc)/int(spbtline)), 1])
        pgnumcol = 30
        pcnt = 0
        pagenum = 0
        ########################################################################
        cv.setFont(fn,fz2)
        lastline = line
        itemspc = sum([self.pght, -sum([usedspc, btmspc])])
        pagemaxline = sum([(int(itemspc)/int(spbtline))])
        pgbghdline = sum([lastline,(pagemaxline*2), 20])
        totpage = (len(itmdct)-1)/pagemaxline
        for r in range(len(itmdct)):
            #cv.setFont(fn,fz2)
            lastline = sum([lastline, -spbtline]) ### Adding Next Line For Next Item
            for k, v in rdct.iteritems():
                try:
                    cv.drawString(rdct[k][0], lastline, itmdct[idx][k])
                except:
                    pass
            ##############################################
            pcnt += 1
            if pcnt%pagemaxline == pagemaxline-1:
                pagenum += 1
                cv.setFont(fn,fz2)
                cv.drawString(self.pgwt-100, pgnumcol, 'Page No:%s/%s'%(str(pagenum), str(totpage)))
                ### New Page Start, Giving Pre-Defined PAGE HEIGHT
                ### And Top Page Padding/Margin, then start write text line 
                lastline = self.pght-55 
                ### Increaing Number of Line After First Page 
                pagemaxline = sum([pagemaxline, sum([pagemaxline, 5])])
                cv.showPage()
                cv.setFont(fn,fz2)
            ##############################################
            line -= spbtline
            idx += 1
        #cv.setFont(fn,fz2)
        cv.drawString(self.pgwt-100, pgnumcol, 'Page :End')
        
    def VCV_MLPG(self, cv, d, fn, fz, ds, lr, lc):
        try:
            cv.setFont(fn,fz)
            txd = ' '.join([d.get(ds)[0][0], d.get(ds)[1]])
            cv.drawString(lr, lc, txd)
        except:
            pass
    def BOXVCV_MLPG(self, cv, d, fn, fz, ds, lr, lc):
        blc = lc-10
        blr = lr-10
        try:
            wh = d.get(ds)[0][0].split('*')
            w = float(wh[0])*10
            h = float(wh[1])*10
            w, h = self.ValSplit( d.get(ds)[0][0], )
            #reccol = d.get(ds)[0][1][0]-5.0
            #rcrow = d.get(ds)[0][1][1]-2.5
            #############################
            cv.setFillColorRGB(0.95,0.95,0.95)
            ##cv.rect(rectCOLpos,rectROWpos,'rect_width',rect_height, fill=0)##stroke=0 >> 0=NoBorder, 1=Border fill=0 >>0=None,1=Yes
            cv.rect(blr,blc,w,h, stroke=0, fill=1)
            cv.setFillColorRGB(0,0,0)
            #############################
        except (ValueError, IndexError):
            pass
        try:
            cv.setFont(fn,fz)
            #txd = ' '.join([d.get(ds)[0][0], d.get(ds)[1]])
            txd = d.get(ds)[1]
            cv.drawString(lr,lc, txd)
        except:
            pass

    def BTXCV_MLPG(self, cv, d, fn, fz, txhdct, itmdct, taxmpkey, txsmhd, spbtline, line):
        ##spbtline = 12 ## Default Given
        
        cv.setFont(fn,fz)

        fd, rdct, qtycount, mvalst, gtl  = ITMTAXCOLLECT(d, txhdct, itmdct, taxmpkey, txsmhd, self.igstbool)
        
        for mval in mvalst:
            if self.compval == False:
                for k, v in rdct.iteritems():
                    try:
                        cv.drawString(rdct[k][0], line, mval[k])
                    except:
                        pass
                line -= spbtline
            else:
                cv.drawString(20, line, "Composition taxable not eligible to collect taxes on supplies.")
    
        mtaxr = [28.0, 18.0, 12.0, 5.0,]
        ### Presetting 0.00 to each tax column
        if self.compval == False:
            for i in (set(mtaxr).difference(fd.keys())):
                mval = {'BTXC':str(i), 'BIGT':'0.00', 'BCGT':'0.00', 'BSGT':'0.00', 'TAXP':'0.00', 'BDIS':'0.00', 'BAMT':'0.00', 'BTOT':'0.00'}
                for k, v in rdct.iteritems():
                    cv.drawString(rdct[k][0], line, mval[k])
                line -= spbtline
        
        if self.compval == False:
            ## gtotlst = [sumigst, sumcgst, sumsgst, totaxpd, totdisc, totamt, gttotl, gttot_adj, gttotlall, gttotl]
            mval = {'BTXC':'Total', 'BIGT':gtl[0], 'BCGT':gtl[1], 'BSGT':gtl[2], 'TAXP':gtl[3], 'BDIS':gtl[4], 'BAMT':gtl[5], 'BTOT':gtl[6]}
            for k, v in rdct.iteritems():
                cv.drawString(rdct[k][0], line, mval[k])
        adjval = Round_Amt_Adj(float(gtl[9]), float(gtl[6]))
        btlst = [str(len(itmdct)), format(sum(qtycount), '0.0f'), gtl[4], gtl[5], gtl[9], adjval]
        
        return btlst
    
    def ITMDCV_default(self, cv, d, fn, fz2, itm, line, icol, sitmlst, itmdct, itmhdct, spbtline):
        rdct = {i.keys()[0] : i.values()[0][0][1] for i in itmhdct}
        ihk = itmdct.values()[0].keys()
        idx = 0
        cv.setFont(fn,fz2)
        for r in range(len(itmdct)):
            for k, v in rdct.iteritems():
                if not isinstance(rdct[k][0], str):
                    ### if any column removed by user, will be excluded in this iteration
                    try:
                        cv.drawString(rdct[k][0], line, itmdct[idx][k])
                    except :
                        pass
            line -= spbtline
            idx += 1
            
    def ITMDCV_iteminfo(self, cv, d, fn, fz2, itm, line, icol, sitmlst, itmdct, itmhdct, spbtline):
        rdct = {i.keys()[0] : i.values()[0][0][1] for i in itmhdct}
        
        ##ihk = itmdct.values()[0].keys()
        idx = 0
        itemline = 0
        cv.setFont(fn,fz2)
        itmrelfnt = sum([int(float(fz2/2)), 1])
        itemspbtline = sum([spbtline/2, 1]) ### *** MUST Calculate Above spbtline sum
        spbtline = sum([spbtline, spbtline/2])
        for r in range(len(itmdct)):
            for k, v in rdct.iteritems():
                if not isinstance(rdct[k][0], str):
                    ### if any column removed by user, will be excluded in this iteration
                    if k == 'ITMN':
                        itms = itmdct[idx]['ITMN'].split('\n')
                        itemline = line 
                        cv.drawString(rdct['ITMN'][0], itemline, itms[0])
                        itemline -= itemspbtline
                        cv.setFont(fn,itmrelfnt)
                        cv.drawString(rdct['ITMN'][0], itemline, itms[1])
                        cv.setFont(fn,fz2)
                    else:
                        try:
                            cv.drawString(rdct[k][0], line, itmdct[idx][k])
                        except :
                            pass
            
            line -= spbtline
            
            idx += 1
            
    def BTXCV(self, cv, d, fn, fz, txhdct, itmdct, taxmpkey, txsmhd, spbtline):
        ##spbtline = 12 ## Default Given
        try:
            line = sum([d.get('BTXC')[0][1][1], -15.0])
        except :
            return
        cv.setFont(fn,fz)

        fd, rdct, qtycount, mvalst, gtl  = ITMTAXCOLLECT(d, txhdct, itmdct, taxmpkey, txsmhd, self.igstbool)
        
        for mval in mvalst:
            if self.compval == False:
                for k, v in rdct.iteritems():
                    try:
                        cv.drawString(rdct[k][0], line, mval[k])
                    except:
                        pass
                line -= spbtline
            else:
                cv.drawString(20, line, "Composition taxable not eligible to collect taxes on supplies.")
    
        mtaxr = [28.0, 18.0, 12.0, 5.0,]
        ### Presetting 0.00 to each tax column
        if self.compval == False:
            for i in (set(mtaxr).difference(fd.keys())):
                mval = {'BTXC':str(i), 'BIGT':'0.00', 'BCGT':'0.00', 'BSGT':'0.00', 'TAXP':'0.00', 'BDIS':'0.00', 'BAMT':'0.00', 'BTOT':'0.00'}
                ###for k, v in rdct.iteritems():
                ###    cv.drawString(rdct[k][0], line, mval[k])
                ### NOT Print those Taxes which are not INCLUDED in Current Bill, ONLY Need Blank Space, Which is REQUIRED 
                #line -= spbtline
        
        if self.compval == False:
            ## gtotlst = [sumigst, sumcgst, sumsgst, totaxpd, totdisc, totamt, gttotl, gttot_adj, gttotlall, gttotl]
            mval = {'BTXC':'Total', 'BIGT':gtl[0], 'BCGT':gtl[1], 'BSGT':gtl[2], 'TAXP':gtl[3], 'BDIS':gtl[4], 'BAMT':gtl[5], 'BTOT':gtl[6]}
            for k, v in rdct.iteritems():
                cv.drawString(rdct[k][0], line, mval[k])
        adjval = Round_Amt_Adj(float(gtl[9]), float(gtl[6]))
        btlst = [str(len(itmdct)), format(sum(qtycount), '0.0f'), gtl[4], gtl[5], gtl[9], adjval]
        
        return btlst

def Round_Amt_Adj(rundval, flotval):
    try:
        adjval = str(abs(flotval-rundval))
    except :
        adjval = '0'
    return adjval
    

def JoinVal(rdic, key):
    try:
        return ':'.join([key, rdic[key][0]]) 
    except (TypeError, KeyError):
        return ':'.join([key, rdic[key][0][0]])
    
class PageSetup1(wx.Panel):
    def __init__(self, parent, getdsp_size, resource_dic):
        self.resource_dic = resource_dic
        self.parent = parent
        wx.Panel.__init__(self, parent)
        self.itmlst1 = [[u'3004',u'PANTOP D CAP', '10*50','20', '10+10', '85.80', '22.222', '17451.6', '12.0','10.71','2344.0','08/19', '888.23', 'ASZ123' ],
            [u'3004',u'OTEK AC PLUS E/D', '15*2','10', '0', '50.0', '22.222', '4550.0', '18.0','10.71','8568.0','08/18', '676.23','ASZ123'],
            [u'2768',u'SAMPLE ITEM SECOND','1Kit', '25', '10+10', '150.0', '22.222', '14500.0', '5.0','10.71','2288.0', '03/19', '675.23','ASZ123'],
            [u'3004',u'ACILOC 150 TAB','10*10', '12', '10+5', '21.7', '22.222', '4453.4', '12.0','10.71','8865.0','01/20','343.23','ASZ123'],
            [u'1234',u'THIRD SAMPLE ITEM','1UNIT', '15', '10+5', '21.7', '22.222', '32445.5', '5.0','10.71', '9786.0', '01/20', '764.23','ASZ123'],
            [u'3004',u'ALTHROCIN 250 TAB','1*10', '14', '10+10', '36.0', '22.222', '7442.0', '18.0','10.71', '8954.0', '04/19', '321.23','ASZ123'],
            [u'3004',u'SOME SAMPLE ITEMS','PACK', '2', '10', '36.0', '22.222', '7442.0', '12.0','10.71', '464.0', '04/19', '443.23','ASZ123'],
            [u'3004',u'ALTHROCIN SYP','200ML', '11', '10+1', '36.10', '22.222', '7442.0',  '28.0','10.71', '1246.0', '04/19', '896.23','ASZ123'],
            [u'2344',u'ANOTHER SAMPLE BAG','10*10', '2', '0', '21.78', '22.222', '3443.4', '12.0','10.71', '3242.0', '01/20', '888.23','ASZ123'],
            [u'3004',u'ANOTHER SAMPLE SMALL','1PCS', '25', '0', '150.0', '22.222', '14500.0', '5.0','10.71','2288.0', '', '675.23','ASZ123'],
            [u'3004',u'ALTHROCIN 250 TUBE','OINT', '2', '10', '36.0', '22.222', '7442.0', '12.0','10.71', '464.0', '04/19', '443.23','ASZ123'],
            [u'2304',u'SAMPLE ITEM TEST','10*10', '2', '10+1', '21.77', '22.222', '3443.4', '18.0','10.71', '3425.0', '01/20', '888.23','ASZ123'],
            [u'3004',u'ACILOC 150 TAB','10*10', '15', '10+10', '21.7', '22.222', '32445.5', '12.0','10.71', '8975.0', '01/20', '555.23','ASZ123'],
            [u'3004',u'ALTHROCIN SYP','200ML', '11', '10+1', '36.10', '22.222', '7442.0',  '28.0','10.71', '1246.0', '04/19', '896.23','ASZ123'],
            [u'2344',u'LAST ITEM','1BOX', '2', '0', '21.78', '22.222', '3443.4', '12.0','10.71', '3242.0', '01/20', '888.23','ASZ123'],]
        self.grid = wx.grid.Grid(self, wx.ID_ANY, size=(1200, 405))
        self.PGPATH1, self.PGPATH2, self.PGPATH3 = Get_pgsett_Path()
        self.PDC = PRINT_DISCRIPTION() 
        self.mcol = 102
        crow = 102 ## 102 * 10
        self.CreateMgrid(crow, self.mcol)
        
        pgwd, pght = '600', '%s'%(crow*10)
        fonts = [7, 8, 9, 9, 9, 15, 20, 7, 8, 9]
        self.prnt = wx.Button(self, -1, "PRINT TEST", (10,210))
        self.prnt.Bind(wx.EVT_BUTTON, self.OnPrint)
        msg = 'This is Sample PrintOut, \nSee and plan to customize \nyour own print page format \n '
        RMS_ToolTip(self.prnt, msg, htx='Sample PrintOut', ftx='',hl=True,fl=False, dly=10)
        
        self.close = wx.Button(self, -1, "Close", (10,210))
        self.close.Bind(wx.EVT_BUTTON, self.OnClose)

        self.loadindpg4 = wx.Button(self, -1, "Load \nSeperate A4", (700,550))
        self.loadindpg4.Bind(wx.EVT_BUTTON, self.OnLoadindpg4)
        self.saveindpg4 = wx.Button(self, -1, "Save \nSeperate A4", (600,550))
        self.saveindpg4.Bind(wx.EVT_BUTTON, self.OnSaveindpg4)
        
        self.landpgpath4 = rmss_config.my_icon(self.PGPATH3)
        if not os.path.exists(rmss_config.my_icon(self.landpgpath4)):
            self.loadindpg4.Disable()
            self.saveindpg4.Disable()
            
        self.Create_indpg4 = wx.Button(self, -1, "Create \nSeperate A4", (800,550))
        self.Create_indpg4.Bind(wx.EVT_BUTTON, self.OnCreate_indpg4)
        
        self.save = wx.Button(self, -1, "Reg Save", )
        self.save.Bind(wx.EVT_BUTTON, self.OnSave)
        msg = "Save PrintOut Format For, \nTAX INVOICE, Regular TAXPAYER's \n "
        RMS_ToolTip(self.save, msg, htx='Regular TaxPayer Save', ftx='',hl=True,fl=False, dly=10)
        self.compsave = wx.Button(self, -1, "Comp Save", )
        self.compsave.Bind(wx.EVT_BUTTON, self.OnCompSave)
        msg = "Save PrintOut Format For, \nBill Of Supply, Composition TAXPAYER's \n "
        RMS_ToolTip(self.compsave, msg, htx='Composition TaxPayer Save', ftx='',hl=True,fl=False, dly=10)
        self.load = wx.Button(self, -1, "Load Settings", )
        self.load.Bind(wx.EVT_BUTTON, self.OnLoad)
        msg = "Load Saved Print Format,\nof TAX INVOICE \nRegular TAXPAYER's \n "
        RMS_ToolTip(self.load, msg, htx='Load GST INVOICE Format', ftx='',hl=True,fl=False, dly=10)
        self.compload = wx.Button(self, -1, "Load Comp Sett.", )
        msg = "Load Saved Print Format of,\nBill Of Supply Composition TAXPAYER's \n "
        RMS_ToolTip(self.compload, msg, htx='Load Bill Of Supply Format', ftx='',hl=True,fl=False, dly=10)
        self.compload.Bind(wx.EVT_BUTTON, self.OnCompLoad)
        self.save.Disable()
        self.compsave.Disable()
    
        self.lps = wx.StaticText(self, -1, "\tLoad Page Setting Values >>> ", size=(200,25))
        self.a4 = wx.RadioButton(self, wx.ID_ANY, "A4 Size", (65, 40), style=wx.ALIGN_RIGHT)#|wx.RB_GROUP)
        self.ahalf = wx.RadioButton(self, wx.ID_ANY, "A4 Half Page", (85, 40), style=wx.ALIGN_RIGHT)#|wx.RB_GROUP)
        self.a2 = wx.RadioButton(self, wx.ID_ANY, "8cm Width", (65, 40), style=wx.ALIGN_RIGHT)#|wx.RB_GROUP)
        self.a1 = wx.RadioButton(self, wx.ID_ANY, "6cm Width", (65, 40), style=wx.ALIGN_RIGHT)#|wx.RB_GROUP)
        self.a0 = wx.RadioButton(self, wx.ID_ANY, "Small Size", (65, 40), style=wx.ALIGN_RIGHT)#|wx.RB_GROUP)

        self.a4.Bind(wx.EVT_RADIOBUTTON, self.Ona4)
        self.ahalf.Bind(wx.EVT_RADIOBUTTON, self.Onahalf)
        self.a2.Bind(wx.EVT_RADIOBUTTON, self.Ona2)
        self.a1.Bind(wx.EVT_RADIOBUTTON, self.Ona1)
        self.a0.Bind(wx.EVT_RADIOBUTTON, self.Ona0)

        self.srdwnstx = wx.StaticText (self, -1, '| Shift Row Down:', )
        self.srdwntx = Rmss_TextCtrl_Num(self, -1, '1', size=(30, 20), validator = MyValidator(FLOAT_ONLY), style=wx.TE_PROCESS_ENTER)
        
        msg = "Select Row Number, If value\nis NOT EMPTY in next cells \nSelected Row Number Value \nwill NOT Move further \n "
        RMS_ToolTip(self.srdwntx, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        self.srdwntx.Bind(wx.EVT_TEXT_ENTER, self.Onsrdwntx)

        self.srupstx = wx.StaticText (self, -1, '| Shift Row UP:', )
        self.sruptx = Rmss_TextCtrl_Num(self, -1, '2', size=(30, 20), validator = MyValidator(FLOAT_ONLY), style=wx.TE_PROCESS_ENTER)
        
        msg = "Select Row Number, If value\nis NOT EMPTY in next cells \nSelected Row Number Value \nwill NOT Move further \n "
        RMS_ToolTip(self.sruptx, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        self.sruptx.Bind(wx.EVT_TEXT_ENTER, self.Onsruptx) 

        self.sclftstx = wx.StaticText (self, -1, '| Shift Col Left:', )
        self.sclfttx = Rmss_TextCtrl_Num(self, -1, '1', size=(30, 20), validator = MyValidator(FLOAT_ONLY), style=wx.TE_PROCESS_ENTER)
        msg = "Select Col Number \nThat Column Has to Shifted \n "
        RMS_ToolTip(self.sclfttx, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.sclfttx1 = Rmss_TextCtrl_Num(self, -1, '8.27', size=(30, 20), validator = MyValidator(FLOAT_ONLY), style=wx.TE_PROCESS_ENTER)
        msg = "Select Column Range \nExample 8.27 Means,\nFrom Row 8 to 27 \n "
        RMS_ToolTip(self.sclfttx1, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.sclfttx.Bind(wx.EVT_TEXT_ENTER, self.ShiftColLeft)
        self.sclfttx1.Bind(wx.EVT_TEXT_ENTER, self.ShiftColLeft1)

        self.sclftrstx = wx.StaticText (self, -1, '| Shift Col Right:', )
        self.sclftrtx = Rmss_TextCtrl_Num(self, -1, '1', size=(30, 20), validator = MyValidator(FLOAT_ONLY), style=wx.TE_PROCESS_ENTER)
        msg = "Select Col Number \nThat Column Has to Shifted \n "
        RMS_ToolTip(self.sclftrtx, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.sclftrtx1 = Rmss_TextCtrl_Num(self, -1, '8.27', size=(30, 20), validator = MyValidator(FLOAT_ONLY), style=wx.TE_PROCESS_ENTER)
        msg = "Select Column Range \nExample 8.27 Means,\nFrom Row 8 to 27 \n "
        RMS_ToolTip(self.sclftrtx1, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        self.sclftrtx.Bind(wx.EVT_TEXT_ENTER, self.ShiftColRight)
        self.sclftrtx1.Bind(wx.EVT_TEXT_ENTER, self.ShiftColRight1)
        
        self.fh1nm = 'Courier-Bold' 
        self.fh2nm = 'Times-Bold'
        self.fh3nm = 'Helvetica'
        self.fnptnm = 'Helvetica-Bold'
        self.fitmnm = "Courier"
        self.fh4nm = "Times-Roman"
        
        self.pswtx = wx.StaticText (self, -1, 'Page Width', )
        self.pwtx = Rmss_TextCtrl_Num(self, -1, pgwd, size=(30, 25), validator = MyValidator(FLOAT_ONLY))
        self.hsttx = wx.StaticText (self, -1, 'Page Height', (220,210))
        self.httx = Rmss_TextCtrl_Num(self, -1, pght, size=(30, 25), validator = MyValidator(FLOAT_ONLY))
        
        self.fsh1 = wx.StaticText (self, -1, 'Head1 Font', (345,210))
        msg = "Font Size Used in \nOwner Name Display,\nNamed as fh1 in RMS \n "
        RMS_ToolTip(self.fsh1, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        self.fh1 = Rmss_TextCtrl_Num(self, -1, str(fonts[6]), size=(30, 25), validator = MyValidator(FLOAT_ONLY))
        msg = "Font Size Used in \nOwner Name Display,\nNamed as fh1 in RMS \n "
        RMS_ToolTip(self.fh1, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fh1n = wx.TextCtrl (self, -1, self.fh4nm,  size=(65,25), style=wx.TE_READONLY)
        msg = "Font Name Used in \nOwner Name Display,\nNamed as fh1n in RMS \n "
        RMS_ToolTip(self.fh1n, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fsh2 = wx.StaticText (self, -1, 'Head2 Font', (435,210))
        msg = "Font Size Used in \nParty Name Display,\nNamed as fh2 in RMS \n "
        RMS_ToolTip(self.fsh2, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fh2 = Rmss_TextCtrl_Num(self, -1, str(fonts[5]), size=(30, 25), validator = MyValidator(FLOAT_ONLY))
        msg = "Font Size Used in \nParty Name Display,\nNamed as fh2 in RMS \n "
        RMS_ToolTip(self.fh2, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        self.fh2n = wx.TextCtrl (self, -1, self.fh4nm,  size=(65,25), style=wx.TE_READONLY)
        msg = "Font Name Used in \nParty Name Display,\nNamed as fh2n in RMS \n "
        RMS_ToolTip(self.fh2n, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fsh3 = wx.StaticText (self, -1, 'Head3 Font', (535,210))
        msg = "Font Size Used in \nBillNo,BillDate,DiscountAmount,TotalAmount,OtherSum,\nNamed as fh3 in RMS \n "
        RMS_ToolTip(self.fsh3, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fh3 = Rmss_TextCtrl_Num(self, -1, str(fonts[4]), size=(30, 25), validator = MyValidator(FLOAT_ONLY))
        msg = "Font Size Used in \nBillNo,BillDate,DiscountAmount,TotalAmount,OtherSum,\nNamed as fh3 in RMS \n "
        RMS_ToolTip(self.fh3, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fh3n = wx.TextCtrl (self, -1, self.fh4nm, size=(65,25), style=wx.TE_READONLY)#fh1nm
        msg = "Font Name Used in \nBillNo,BillDate,DiscountAmount,TotalAmount,OtherSum,\nNamed as fs1n in RMS \n "
        RMS_ToolTip(self.fh3n, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fsitmh = wx.StaticText (self, -1, 'ItemHead Font', (635 ,210))
        msg = "Font Size Used in \nItems Heading, Authorized Signatory,\nNamed as fitmh in RMS \n "
        RMS_ToolTip(self.fsitmh, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fitmh = Rmss_TextCtrl_Num(self, -1, str(fonts[3]), size=(30, 25), validator = MyValidator(FLOAT_ONLY))
        msg = "Font Size Used in \nItems Heading, Authorized Signatory,\nNamed as fitmh in RMS \n "
        RMS_ToolTip(self.fsitmh, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fitmnh = wx.TextCtrl (self, -1, self.fh4nm,  size=(65,25), style=wx.TE_READONLY)
        msg = "Font Name Used in \nItems Heading, Authorized Signatory,\nNamed as fitnh in RMS \n "
        RMS_ToolTip(self.fitmnh, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fsitm = wx.StaticText (self, -1, 'Item Font', (700 ,210))
        msg = "Font Size Used in \nItems Text,\nNamed as fitm in RMS \n "
        RMS_ToolTip(self.fsitm, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fitm = Rmss_TextCtrl_Num(self, -1, str(fonts[2]), size=(30, 25), validator = MyValidator(FLOAT_ONLY))
        msg = "Font Size Used in \nItems Text,\nNamed as fitm in RMS "
        RMS_ToolTip(self.fitm, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fitmn = wx.TextCtrl (self, -1, self.fh4nm,  size=(65,25), style=wx.TE_READONLY)
        msg = "Font Size Used in \nGST TAX Summary Page Bottom,\nNamed as fh6 in RMS \n "
        RMS_ToolTip(self.fitmn, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fsh4 = wx.StaticText (self, -1, 'Head4 Font', )
        msg = "Font Size Used in \nCash-Credit, OwnerDetails, ItemCount, QuantityCount,\nAmount in Words, Comments, Authorised Owner,\nNamed as fh4 in RMS \n "
        RMS_ToolTip(self.fsh4, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fh4 = Rmss_TextCtrl_Num(self, -1, str(fonts[1]), size=(30, 25), validator = MyValidator(FLOAT_ONLY))
        msg = "Font Size Used in \nCash-Credit, OwnerDetails, ItemCount, QuantityCount,\nAmount in Words, Comments, Authorised Owner,\nNamed as fh4 in RMS \n "
        RMS_ToolTip(self.fh4, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fh4n = wx.TextCtrl (self, -1, self.fh4nm,  size=(65,25), style=wx.TE_READONLY)
        msg = "Font Name Used in \nCash-Credit, OwnerDetails, ItemCount, QuantityCount,\nAmount in Words, Comments, Authorised Owner,\nNamed as fh4n in RMS \n "
        RMS_ToolTip(self.fh4n, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fsh5 = wx.StaticText (self, -1, 'Head5 Font', )
        msg = "Font Size Used in \nDisclaimers at Page Bottom,\nNamed as fh5 in RMS \n "
        RMS_ToolTip(self.fsh5, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fh5 = Rmss_TextCtrl_Num(self, -1, str(fonts[7]), size=(30, 25), validator = MyValidator(FLOAT_ONLY))
        msg = "Font Size Used in \nDisclaimers at Page Bottom,\nNamed as fh5 in RMS \n "
        RMS_ToolTip(self.fh5, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fh5n = wx.TextCtrl (self, -1, self.fh4nm,  size=(65,25), style=wx.TE_READONLY)
        msg = "Font Name Used in \nDisclaimers at Page Bottom,\nNamed as fh5n in RMS \n "
        RMS_ToolTip(self.fh5n, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fsh6 = wx.StaticText (self, -1, 'TaxSum Font', )
        msg = "Font Size Used in \nGST TAX Summary Page Bottom,\nNamed as fh6 in RMS \n "
        RMS_ToolTip(self.fsh6, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fh6 = Rmss_TextCtrl_Num(self, -1, str(fonts[8]), size=(30, 25), validator = MyValidator(FLOAT_ONLY))
        msg = "Font Size Used in \nGST TAX Summary Page Bottom,\nNamed as fh6 in RMS \n "
        RMS_ToolTip(self.fh6, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fh6n = wx.TextCtrl (self, -1, self.fh4nm,  size=(65,25), style=wx.TE_READONLY)
        msg = "Font Name Used in \nGST TAX Summary Page Bottom,\nNamed as fh6n in RMS \n "
        RMS_ToolTip(self.fh6n, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fsn2pt = wx.StaticText (self, -1, 'Party Font2', )
        msg = "Font Size Used in \nParty Details Page Top,\nNamed as f2pt in RMS \n "
        RMS_ToolTip(self.fsn2pt, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fn2pt = Rmss_TextCtrl_Num(self, -1, str(fonts[9]), size=(30, 25), validator = MyValidator(FLOAT_ONLY))
        msg = "Font Size Used in \nParty Details Page Top,\nNamed as f2pt in RMS \n "
        RMS_ToolTip(self.fn2pt, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fn2ptn = wx.TextCtrl (self, -1, self.fh4nm,  size=(65,25), style=wx.TE_READONLY)
        msg = "Font Name Used in \nParty Details Page Top,\nNamed as f2ptn in RMS \n "
        RMS_ToolTip(self.fn2ptn, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fsns = wx.StaticText (self, -1, 'Small Font', (460,240))
        msg = "Font Size Used in \nTop Greating Message Page Top,\nNamed as fs1 in RMS \n "
        RMS_ToolTip(self.fsns, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fns = Rmss_TextCtrl_Num(self, -1, str(fonts[0]), size=(30, 25), validator = MyValidator(FLOAT_ONLY))
        msg = "Font Size Used in \nTop Greating Message Page Top,\nNamed as fs1 in RMS \n "
        RMS_ToolTip(self.fns, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.fnsn = wx.TextCtrl (self, -1, self.fh1nm,  size=(65,25), style=wx.TE_READONLY)
        msg = "Font Name Used in \nTop Greating Message Page Top,\nNamed as fs1n in RMS \n "
        RMS_ToolTip(self.fnsn, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.gtsh = wx.StaticText (self, -1, 'Grand Total Heading:', )
        self.gth = wx.TextCtrl (self, -1, 'Payable', size=(60,25))
        msg = "User can Change this text \neg:Total, Payable, NetPay\n"
        RMS_ToolTip(self.gth, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.spbtsline = wx.StaticText (self, -1, 'Space Between Itmes:', )
        msg = "Gap Given Between Two Items Disaplay "
        RMS_ToolTip(self.spbtsline, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.spbtline = Rmss_TextCtrl_Num(self, -1, '12', size=(30, 25), validator = MyValidator(FLOAT_ONLY))
        msg = "Gap Given Between Two Items Disaplay "
        RMS_ToolTip(self.spbtline, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.slider = wx.Slider(
            self, 2, 2, 1, 2, (30, 60), (90, -1), 
            #wx.SL_HORIZONTAL ##| wx.SL_LABELS 
            )
        self.lanpot = 2  ## 2 ==> for potrait print mode 1 => for landscape
        self.slider.Bind(wx.EVT_SCROLL_CHANGED, self.OnSliderChanged)

        self.rsendline_stx = wx.StaticText (self, -1, 'V.Line Right Side:', )
        msg = "Vertical Line Position \nOn Right Side of Page \n "
        RMS_ToolTip(self.rsendline_stx, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.rsendline = Rmss_TextCtrl_Num(self, -1, '575', size=(30, 25), validator = MyValidator(FLOAT_ONLY))
        msg = "Vertical Line Position \nOn Right Side of Page \n "
        RMS_ToolTip(self.rsendline, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.getrowline_stx = wx.StaticText (self, -1, 'Page Horizontal Lines:', )
        
        self.getrowline = Rmss_TextCtrl_Num(self, -1, str(self.grid.GetNumberRows()), size=(30, 25), validator = MyValidator(FLOAT_ONLY))
        msg = "Toal Horizontal Lines \nOn Front Page \n "
        RMS_ToolTip(self.getrowline, msg, htx='', ftx='',hl=False,fl=False, dly=10)
        
        self.delfiles = wx.Button(self, -1, "Delete My Sett.", (30,550))
        msg = "Delete All User Print Format \n Activate RMS Built-in Print Format\nUser Can Again Customize These Settings\n\n"
        RMS_ToolTip(self.delfiles, msg, htx='Delete User Print Format', ftx='RMS In-Built Print Format Activate',hl=True,fl=True, dly=10)
        self.delfiles.SetForegroundColour(wx.RED)
        
        self.fnnlst = wx.ListCtrl(self, wx.ID_ANY,  style=wx.LC_REPORT|wx.LC_NO_HEADER)
        index = 0
        self.fnnlst.InsertColumn(0, '')
        self.fnnlst.SetColumnWidth(0, 100) 
        for d in [self.fh1nm, self.fh2nm, self.fh3nm, self.fnptnm, self.fitmnm, self.fh4nm, ]:
            self.fnnlst.InsertStringItem(index, str(''))
            self.fnnlst.SetStringItem(index, 0, str(d))
        self.fnnlst.SetInitialSize((110, 70))
        self.Restrdic()
        rdic = sorted(self.rdic.iteritems())
        
        ss = '\n'.join([str(rdic)[:300],str(rdic)[300:600],str(rdic)[600:900],str(rdic)[900:1200],
            str(rdic)[1200:1500],str(rdic)[1500:1700],str(rdic)[1700:2100],str(rdic)[2100:],])
            #str(rdic)[1800:2000], str(rdic)[2000:2200]],)
        #info = wx.StaticText (self, -1, str(ss), (0,270))
        self.info = wx.StaticText (self, -1, str(ss),)
        self.Restlinedic(self.mcol)
        self.MyLayout()
        self.loadmsg = False
        self.delfiles.Bind(wx.EVT_BUTTON, self.DelFiles)
        self.getrowline.Bind(wx.EVT_KILL_FOCUS, self.OnRowLineKillFocus)
        self.getrowline.Bind(wx.EVT_TEXT, self.OnRowLine)
        self.fh1n.Bind(wx.EVT_LEFT_DOWN, self.Onfh1n)
        self.fh2n.Bind(wx.EVT_LEFT_DOWN, self.Onfh2n)
        self.fh3n.Bind(wx.EVT_LEFT_DOWN, self.Onfh3n)
        self.fitmnh.Bind(wx.EVT_LEFT_DOWN, self.Onfitmnh)
        self.fn2ptn.Bind(wx.EVT_LEFT_DOWN, self.Onfn2ptn)
        self.fitmn.Bind(wx.EVT_LEFT_DOWN, self.Onfitmn)
        self.fh4n.Bind(wx.EVT_LEFT_DOWN, self.Onfh4n)
        self.fh5n.Bind(wx.EVT_LEFT_DOWN, self.Onfh5n)
        self.fh6n.Bind(wx.EVT_LEFT_DOWN, self.Onfh6n)
        self.fnsn.Bind(wx.EVT_LEFT_DOWN, self.Onfnsn)
        self.fnnlst.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.Onfnnlst)
        self.fnnlst.Hide()
        self.commnobject = None

    def DelFiles(self, event):
        mess =wx.MessageBox("WANT TO Delete User Print Settings ??\nRMS will use built in A4 "
                            "Print Page Settings,\nYou Can Again Customize Bill Print Page Settings.",
               "Wait !!", wx.YES_NO |wx.NO_DEFAULT|wx.ICON_EXCLAMATION)
        if mess == wx.YES:
            if os.path.isfile(rmss_config.my_icon(self.PGPATH1)):
                os.remove(rmss_config.my_icon(self.PGPATH1))
                
            if os.path.isfile(rmss_config.my_icon(self.PGPATH2)):
                os.remove(rmss_config.my_icon(self.PGPATH2))
        
            self.delfiles.SetLabel('Deleted Ok!')
        event.Skip()
        
    def CreateMgrid(self, grow, gcol):
        ccol = [str(r+1) for r in range(gcol)]
        colwid = 30
        self.grid.CreateGrid(grow, gcol)
        for row in range(gcol):
            self.grid.SetColLabelValue(row, ccol[row])
            self.grid.SetColSize(row, colwid)
        self.grid.SetRowLabelSize(25)
        self.grid.SetColLabelSize(25)

    def OnRowLineKillFocus(self, event):
        if self.getrowline.GetValue().strip() == "":
            self.getrowline.SetValue(str(self.grid.GetNumberRows()))
        event.Skip()
        
    def OnRowLine(self, event):
        try:
            numrows = int(self.getrowline.GetValue())
        except:
            numrows = self.grid.GetNumberRows()
        self.httx.ChangeValue(str(numrows*10))
        event.Skip()

    def LandScapeA4(self):
        self.ClearGrid() #
        row = self.grid.GetNumberRows()
        col = self.grid.GetNumberCols()
        hld, vld = {}, {}
        rdic = {'DISS': [[u'Discount :', [680, 300]], ''], 'DPNO': ['Dispatch No.:', ''], 'HSNC': [[u'HSNC', [260, 510]], ''],
                'OEML': [['', [130, 540]], ''], 'BTXC': [[u'GST%', [20, 310]], ''], 'BAMT': [[u'Amount', [290, 310]], ''],
                'IEXP': [[u'Exp.', [660, 510]], ''], 'ONDP': [['', [20, 590]], 'me'], 'PNDP': [['', [290, 570]], ''],
                '__NAME__': [], 'ITAX': [[u'Tax%', [540, 510]], ''], 'TAXP': [[u'Tax Amt.', [190, 310]], ''],
                'DDAT': ['', ''], 'BINF': ['', ''], 'IPAC': [[u'Pack', [300, 510]], ''], 'IMRP': [[u'MRP', [620, 510]], ''],
                'CSCR': [[u'9*1', [360, 590]], ''], 'THDP': [['', [380, 600]], '____'], 'OREG': [['', [20, 550]], ''],
                'IRAT': [[u'Rate', [420, 510]], ''], 'BIGT': [[u'IGST', [50, 310]], ''], 'PAD1': [['', [290, 560]], ''],
                'PPHN': [['', [290, 550]], ''], 'PAD2': [['', [400, 560]], ''], 'GTOT': [[u'18*2', [630, 250]], ''],
                'OAD1': [['', [20, 570]], ''], 'OAD2': [['', [20, 560]], ''], 'IRA2': [[u'Rate2', [460, 510]], ''],
                'INET': [[u'Net', [700, 510]], ''], 'IQTY': [[u'Qty', [340, 510]], ''], 'BDAT': [[u'Bill Date:', [580, 560]], ''],
                'IBON': [[u'Bonus', [380, 510]], ''], 'BCGT': [[u'CGST', [90, 310]], ''], 'PAGESIZE': [950.0, 620.0],
                'AWRD': [[u'Rs.', [20, 240]], ''], 'DPTO': ['Dispatched To:', ''], 'IBAT': [[u'Batch', [740, 510]], ''],
                'AUSS': [[u'Authorized Signatory', [450, 250]], ''], 'ADTS': ['Add Taxes', ''], 'HLIN': ['', 'horizontal line'],
                'CMNT': [[u'MSG :', [400, 310]], ''], 'OGST': [['', [20, 540]], ''], 'DITM': ['Deliver ITems:', ''],
                'IAMT': [[u'Amount', [500, 510]], ''], 'OTH1': [['', [680, 280]], ''], 'QTTC': [[u'Total Qty:', [710, 310]], ''],
                'OTH3': ['', ''], 'OTH2': [['', [680, 270]], ''], 'ORDN': [[u'Order No.:', [580, 550]], ''], 'STA4': ['', ''],
                'VLIN': ['', 'vertical line'], 'STA2': [['', [400, 290]], ''], 'STA3': [['', [400, 280]], ''],
                'LOGO': [[u'5*5', [220, 540]], ''], 'ITMC': [[u'Items:', [680, 310]], ''], 'PMOD': [[u'Pay.Mode:', [580, 540]], ''],
                'BILN': [[u'Bill No:', [580, 570]], ''], 'PREG': [['', [400, 550]], ''], 'IDIS': [[u'Dis%', [580, 510]], ''],
                'AMTS': [[u'Amount :', [680, 290]], ''], 'STA1': [['', [400, 300]], ''], 'AUSF': [[u'For:', [420, 270]], ''],
                'PGST': [[u'GSTN:', [290, 540]], ''], 'OPHN': [['', [100, 560]], ''], 'BSGT': [[u'SGST', [140, 310]], ''],
                'AUSO': [['', [450, 270]], ''], 'BDIS': [[u'Discount', [240, 310]], ''], 'BTOT': [[u'Total', [340, 310]], ''],
                'ITMN': [[u'ITEM NAME', [30, 510]], ''], 'PEML': [['', [400, 540]], ''], 'AUTS': ['', '']}
        for i in range(1, row):
            hld[''.join(['_HLS', '%02d'%i])] = ['', '', '']
            hld[''.join(['_HLE', '%02d'%i])] = ['', '', '']
            
        hld['LPRT'] = [1,1,810]
        hldrel = hld['LPRT'][2]
        
        hld['_HLE02'] = [770, 520, '']
        hld['_HLE03'] = [770, 510, '']
        hld['_HLE06'] = [770, 310, '']
        hld['_HLE04'] = [770, 320, '']
        hld['_HLS04'] = [10, 320, '']
        hld['_HLS06'] = [10, 310, '']
        hld['_HLS02'] = [10, 520, '']
        hld['_HLS03'] = [10, 510, '']
        
        for i in range(1, row):
            vld[''.join(['_VLS', '%02d'%i])] = ['', '', '']
            vld[''.join(['_VLE', '%02d'%i])] = ['', '', '']
        vld['_VLE04'] = [300, 320, '']
        vld['_VLE05'] = [340, 320, '']
        vld['_VLE06'] = [380, 320, '']
        vld['_VLE07'] = [420, 320, '']
        vld['_VLE03'] = [260, 320, '']
        vld['_VLE08'] = [460, 320, '']
        vld['_VLE09'] = [500, 320, '']
        vld['_VLE13'] = [660, 320, '']
        vld['_VLE12'] = [620, 320, '']
        vld['_VLE10'] = [540, 320, '']
        vld['_VLE14'] = [700, 320, '']
        vld['_VLE11'] = [580, 320, '']
        vld['_VLE16'] = [390, 240, '']
        vld['_VLE15'] = [740, 320, '']
        vld['_VLS03'] = [260, 500, '']
        vld['_VLS06'] = [380, 500, '']
        vld['_VLS07'] = [420, 500, '']
        vld['_VLS04'] = [300, 500, '']
        vld['_VLS05'] = [340, 500, '']
        vld['_VLS08'] = [460, 500, '']
        vld['_VLS09'] = [500, 500, '']
        vld['_END0'] =  [10, 230, '']
        vld['_VLS11'] = [580, 500, '']
        vld['_VLS10'] = [540, 500, '']
        vld['_VLS13'] = [660, 500, '']
        vld['_VLS12'] = [620, 500, '']
        vld['_VLS15'] = [740, 500, '']
        vld['_VLS14'] = [700, 500, '']
        vld['_VLS16'] = [390, 310, '']

        fnt = {'f2pt': 9.0, 'fit': 9.0, 'fitmh': 9.0, 'fs1': 7.0, 'fh4': 8.0, 'fh5':
               7.0, '__name__': 'fonts', 'fh6': 8.0, 'fh1': 18.0, 'fh3': 9.0, 'fh2': 12.0}

        fnm = {'fh6n': 'Times-Roman', 'fh4n': 'Times-Roman', 'fitn': 'Times-Roman',
               'f2ptn': 'Times-Roman', 'fitnh': 'Times-Roman','fh1n': 'Times-Roman',
               '__name__': 'ftname', 'fh5n': 'Times-Roman', 'fs1n': 'Courier-Bold', 'fh3n': 'Times-Roman', 'fh2n': 'Times-Roman'}
        self.spbtline.SetValue(str('10.0'))
        self.gth.SetValue(str('Payable'))
        self.rsendline.ChangeValue(str(hldrel))
        spbtline, gtothead = '10.0', 'Payable'
        self.LoadAll(rdic, hld, vld, fnt, fnm, spbtline, gtothead)
        
    def OnSliderChanged(self, event):
        self.lanpot = int(self.slider.GetValue())
        if self.lanpot == 1:
            self.LandScapeA4()
        else:
            self.Onahalf(event)
            self.rsendline.ChangeValue(str('575'))
        event.Skip()
        
    def MoveRows(self, sdwn, rsdwn, ):
        grows = self.grid.GetNumberRows()
        gcols = self.grid.GetNumberCols()
        sgrdval = []
        for r in (range(rsdwn, grows)):
            ilst = [self.grid.GetCellValue(r, c) for c in range(gcols)]
            sgrdval.append(ilst)
            
        for r in (range(rsdwn, grows)):
            for c in range(self.grid.GetNumberCols()):
                self.grid.SetCellValue(r, c, '')
        return sgrdval
    
    def Onsrdwntx(self, event):
        sdwn = int(self.srdwntx.GetValue())
        rsdwn = sdwn-1
        sgrdval = self.MoveRows(sdwn, rsdwn, )
        for r in (range(len(sgrdval))):
            for c in range(self.grid.GetNumberCols()):
                try:
                    self.grid.SetCellValue(r+sdwn, c, sgrdval[r][c])
                except :
                    pass
        event.Skip()
    def ShiftColRight(self, event):
        self.ShiftColRight1(event)
    def ShiftColRight1(self, event):
        self.ShiftCol_L_R(event, 53, 1, self.sclftrtx1, self.sclftrtx, "R")
        
    def ShiftColLeft(self, event):
        self.ShiftColLeft1(event)
    def ShiftColLeft1(self, event):
        self.ShiftCol_L_R(event, 1, -1, self.sclfttx1, self.sclfttx, "L")
        
    def ShiftCol_L_R(self, event, stcol, addint, rngtxtfld, txtfld, rl):
        grows = self.grid.GetNumberRows()
        rrng = rngtxtfld.GetValue().split('.')
        try:
            frm, to = int(rrng[0])-1, int(rrng[1])
        except:
            return
        try:
            scoll = int(txtfld.GetValue())-1
        except:
            return
        scoll_r = sum([scoll, addint])
        sgrdval = []
        for r in (range(grows)):
            sgrdval.append(self.grid.GetCellValue(r, scoll))
        ftl = []
        for r in range(frm, to):
            if self.grid.GetCellValue( r, scoll_r) == "":
                ## Empty Cell for Shifted Value 
                self.grid.SetCellValue(r, scoll, '')
            else:
                ### Restore Old Value if Shifting Value is Present in Next Shift Cell  
                self.grid.SetCellValue(r, scoll, self.grid.GetCellValue( r, scoll_r))
                event.Skip()
                
        if int(scoll) > 0 :
            for r in range(frm, to):
                self.grid.SetCellValue(r, scoll_r, sgrdval[r])
        sgrdval = []
        
        if rl == "L":
            if int(txtfld.GetValue()) > stcol :
                ### Adjust Text Field Col Number to Next Col Number in Left Side
                txtfld.SetValue(str(int(txtfld.GetValue())+addint))
        else:
            if int(txtfld.GetValue()) < stcol :
                ### Adjust Text Field Col Number to Next Col Number in Right Side
                txtfld.SetValue(str(int(txtfld.GetValue())+addint))
        event.Skip()
        
    def Onsruptx(self, event):
        try:
            if int(self.sruptx.GetValue()) < 2 :
                return
        except :
            return
        sdwn = int(self.sruptx.GetValue())
        rsdwn = sdwn-1
        grows = self.grid.GetNumberRows()
        gcols = self.grid.GetNumberCols()
        nxtsgrdval = [self.grid.GetCellValue(rsdwn-1, c) for c in range(gcols) if self.grid.GetCellValue(rsdwn-1, c).strip() != ""]
        if nxtsgrdval == []:
            sgrdval = [self.grid.GetCellValue(sdwn-1, c) for c in range(gcols)]
            for c in range(gcols):
                self.grid.SetCellValue(sdwn-1, c, '')
            rsdwn = rsdwn-1
            for c in range(gcols):
                try:
                    if self.grid.GetCellValue(rsdwn, c).strip() == "":
                        self.grid.SetCellValue(rsdwn, c, sgrdval[c])
                        self.sruptx.SetValue(str(sdwn-1))
                    else:
                        event.Skip()
                        return
                except:
                    pass
     
        event.Skip()
        
    def SetEventObjectVal(self, wpos, evob):
        self.fnnlst.Refresh()
        self.commnobject = evob
        self.fnnlst.SetPosition((wpos[0], wpos[1]+25))
        self.fnnlst.Show()
        #self.fnnlst.Raise()
        
    def Onfnnlst(self, event):
        x = self.fnnlst.GetFocusedItem()
        text = self.fnnlst.GetItem(x, 0).GetText()
        self.commnobject.SetValue(text)
        self.fnnlst.Hide()
        event.Skip()
        
    def EvtObjVal(self, event):
        wpos = event.GetEventObject().GetPosition()
        self.SetEventObjectVal(wpos, event.GetEventObject())
        event.Skip()
        
    def Onfh1n(self, event):
        self.EvtObjVal(event)
    def Onfh2n(self, event):
        self.EvtObjVal(event)
    def Onfh3n(self, event):
        self.EvtObjVal(event)
    def Onfitmnh(self, event):
        self.EvtObjVal(event)
    def Onfn2ptn(self, event):
        self.EvtObjVal(event)
    def Onfitmn(self, event):
        self.EvtObjVal(event)
    def Onfh4n(self, event):
        self.EvtObjVal(event)
    def Onfh5n(self, event):
        self.EvtObjVal(event)
    def Onfh6n(self, event):
        self.EvtObjVal(event)
    def Onfnsn(self, event):
        self.EvtObjVal(event)
       
    def MyLayout(self):
        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        sizer1h = wx.BoxSizer(wx.HORIZONTAL)
        sizer_2 = wx.BoxSizer(wx.VERTICAL)
        sizer_3 = wx.BoxSizer(wx.HORIZONTAL)
        sizer3v = wx.BoxSizer(wx.VERTICAL)
        sizer_4 = wx.BoxSizer(wx.HORIZONTAL)
        
        sfz = wx.FlexGridSizer(2, 21, 5, 5)
        sizer_1.Add((30, 10), 0, 0, 0)
        
        sizer_1.Add(sizer1h, 0, wx.EXPAND, 0)
        sizer1h.Add((30, 10), 0, 0, 0)
        sizer1h.Add(self.grid,0,0,0)
        ##################
        sfz0 = wx.FlexGridSizer(1, 16, 5, 5)
        sizer_2.Add(sfz0, 0, wx.EXPAND, 0)
        sfz0.Add(self.lps, 0, 0, 0)
        sfz0.Add(self.a4, 0, 0, 0)
        sfz0.Add(self.ahalf, 0, 0, 0)
        sfz0.Add(self.a2, 0, 0, 0)
        sfz0.Add(self.a1, 0, 0, 0)
        sfz0.Add(self.a0, 0, 0, 0)
        sfz0.Add(self.srdwnstx, 0, 0, 0)
        sfz0.Add(self.srdwntx, 0, 0, 0)
        sfz0.Add(self.srupstx, 0, 0, 0)
        sfz0.Add(self.sruptx, 0, 0, 0)
        
        sfz0.Add(self.sclftstx, 0, 0, 0)
        sfz0.Add(self.sclfttx, 0, 0, 0)
        sfz0.Add(self.sclfttx1, 0, 0, 0)

        sfz0.Add(self.sclftrstx, 0, 0, 0)
        sfz0.Add(self.sclftrtx, 0, 0, 0)
        sfz0.Add(self.sclftrtx1, 0, 0, 0)
        ######################
        sizer_2.Add((30, 10), 0, 0, 0)
        sizer_3.Add((30, 20), 0, 0, 0)
        self.info.Hide()
        sizer_3.Add(self.prnt, 0,0,0)
        sizer_3.Add(self.save, 0,0,0)
        sizer_3.Add(self.compsave, 0,0,0)
        sizer_3.Add(self.load, 0,0,0)
        sizer_3.Add(self.compload, 0,0,0)
        sizer_3.Add(self.close, 0, 0, 0)
        sizer_3.Add((40, 0), 0, 0, 0)
        sizer_3.Add(wx.StaticText(self, -1, "Landscape"), 0, 0, 0)
        sizer_3.Add(self.slider, 0, 0, 0)
        sizer_3.Add(wx.StaticText(self, -1, "Potrait"), 0, 0, 0)
        sizer_3.Add(wx.StaticText(self, -1, " \t "), 0, 0, 0)
        sizer_3.Add(self.rsendline_stx, 0, 0, 0)
        sizer_3.Add(self.rsendline, 0, 0, 0)
        sizer_3.Add(wx.StaticText(self, -1, " \t "), 0, 0, 0)
        sizer_3.Add(self.getrowline_stx, 0, 0, 0)
        sizer_3.Add(self.getrowline, 0, 0, 0)
        
        
        sizer_2.Add(sizer_3, 0, wx.EXPAND, 0)
        sizer_2.Add(sizer3v, 0, wx.EXPAND, 0)
        sizer_2.Add(sizer_4, 0, wx.EXPAND, 0)
        ##sizer3v.Add((30, 20), 0, 0, 0)
        sizer_4.Add((30, 20), 0, 0, 0)
        
        sizer_4.Add(sfz, 0, wx.EXPAND, 0)
        
        sfz.Add(self.fsh1, 0, 0, 0)
        sfz.Add(self.fh1, 0, 0, 0)
        sfz.Add(self.fh1n, 0, 0, 0)
        
        sfz.Add(self.fsh2, 0, 0, 0)
        sfz.Add(self.fh2, 0, 0, 0)
        sfz.Add(self.fh2n, 0, 0, 0)
        
        sfz.Add(self.fsh3, 0, 0, 0)
        sfz.Add(self.fh3, 0, 0, 0) 
        sfz.Add(self.fh3n, 0, 0, 0)

        sfz.Add(self.fsitmh, 0, 0, 0)
        sfz.Add(self.fitmh, 0, 0, 0)
        sfz.Add(self.fitmnh, 0, 0, 0)
        
        sfz.Add(self.fsitm, 0, 0, 0)
        sfz.Add(self.fitm, 0, 0, 0)
        sfz.Add(self.fitmn, 0, 0, 0)

        sfz.Add(self.fsh6, 0, 0, 0)
        sfz.Add(self.fh6, 0, 0, 0)
        sfz.Add(self.fh6n, 0, 0, 0)
        
        sfz.Add(self.spbtsline, 0, 0, 0)
        sfz.Add(self.spbtline, 0, 0, 0)
        sfz.Add(wx.StaticText(self, -1, ""), 0, 0, 0)

        sfz.Add(self.fsh4, 0, 0, 0)
        sfz.Add(self.fh4, 0, 0, 0)
        sfz.Add(self.fh4n, 0, 0, 0)

        sfz.Add(self.fsns, 0, 0, 0)
        sfz.Add(self.fns, 0, 0, 0)
        sfz.Add(self.fnsn, 0, 0, 0)

        sfz.Add(self.fsh5, 0, 0, 0)
        sfz.Add(self.fh5, 0, 0, 0)
        sfz.Add(self.fh5n, 0, 0, 0)

        sfz.Add(self.fsn2pt, 0, 0, 0)
        sfz.Add(self.fn2pt, 0, 0, 0)
        sfz.Add(self.fn2ptn, 0, 0, 0)
        #sfz.Add(wx.StaticText(self, -1, ""), 0, 0, 0)
        #sfz.Add(wx.StaticText(self, -1, ""), 0, 0, 0)
        #sfz.Add(wx.StaticText(self, -1, ""), 0, 0, 0)
        
        sfz.Add(self.pswtx, 0, 0, 0)
        sfz.Add(self.pwtx, 0, 0, 0)
        sfz.Add(wx.StaticText(self, -1, ""), 0, 0, 0)
        
        sfz.Add(self.hsttx, 0, 0, 0)
        sfz.Add(self.httx, 0, 0, 0)
        sfz.Add(wx.StaticText(self, -1, ""), 0, 0, 0)

        sfz.Add(self.gtsh, 0, 0, 0)
        sfz.Add(self.gth, 0, 0, 0)
        sfz.Add(wx.StaticText(self, -1, ""), 0, 0, 0)

        sizer_1.Add(sizer_2, 0, wx.EXPAND, 0)
        #sizer_1.Add((150, 150), 0, 0, 0)
        self.SetSizer(sizer_1)
        sizer_1.Fit(self)
        self.Layout()
        
    def ClearGrid(self):
        self.save.Enable()
        self.compsave.Enable()
        for r in range(self.grid.GetNumberRows()):
            for c in range(self.grid.GetNumberCols()):
                self.grid.SetCellValue(r, c, '')
    def Onahalf(self, event):
        self.ClearGrid()
        self.SetDefaultValue_Half()
        self.loadmsg = False
        event.Skip()
    def Ona4(self, event):
        self.ClearGrid()
        #print self.rdic
        self.SetDefaultValue()
        self.loadmsg = False
        event.Skip()
    def Ona2(self, event):
        self.ClearGrid()
        ### Not Prepare Yet
        self.SetDefaultSize8cm()
        self.loadmsg = False
        event.Skip()
    def Ona1(self, event):
        self.ClearGrid()
        ### Not Prepare Yet
        self.loadmsg = False
        event.Skip()
    def Ona0(self, event):
        self.ClearGrid()
        ### Not Prepare Yet
        self.loadmsg = False
        event.Skip()
        
    def SetDefaultValue(self):
        self.grid.SetCellValue(3, 27, JoinVal(self.rdic, 'THDP'))
        self.grid.SetCellValue(4, 26, JoinVal(self.rdic, 'CSCR'))

        self.grid.SetCellValue(8, 0, '_HLS01')
        self.grid.SetCellValue(8, 54, '_HLE01')
        
        self.grid.SetCellValue(6, 17, JoinVal(self.rdic, 'ONDP'))
        self.grid.SetCellValue(7, 15, JoinVal(self.rdic, 'OAD1'))
        self.grid.SetCellValue(7, 25, JoinVal(self.rdic, 'OAD2'))
        self.grid.SetCellValue(7, 35, JoinVal(self.rdic, 'OEML'))
        self.grid.SetCellValue(8, 15, JoinVal(self.rdic, 'OPHN'))
        self.grid.SetCellValue(8, 18, JoinVal(self.rdic, 'OREG'))
        self.grid.SetCellValue(8, 35, JoinVal(self.rdic, 'OGST'))
        ######################################
        self.grid.SetCellValue(10, 1, JoinVal(self.rdic, 'PNDP'))
        self.grid.SetCellValue(11, 1, JoinVal(self.rdic, 'PAD1'))
        self.grid.SetCellValue(12, 1, JoinVal(self.rdic, 'PREG'))
        self.grid.SetCellValue(13, 1, JoinVal(self.rdic, 'PGST'))

        self.grid.SetCellValue(11, 10, JoinVal(self.rdic, 'PAD2'))
        self.grid.SetCellValue(12, 10, JoinVal(self.rdic, 'PPHN'))
        self.grid.SetCellValue(13, 12, JoinVal(self.rdic, 'PEML'))
        
        self.grid.SetCellValue(10, 26, JoinVal(self.rdic, 'BILN'))
        self.grid.SetCellValue(11, 26, JoinVal(self.rdic, 'PMOD'))
        self.grid.SetCellValue(12, 26, JoinVal(self.rdic, 'DITM'))
        self.grid.SetCellValue(13, 26, JoinVal(self.rdic, 'ORDN'))

        self.grid.SetCellValue(9, 26, '_VLS01')
        self.grid.SetCellValue(14, 26, '_VLE01')
        
        self.grid.SetCellValue(10, 36, JoinVal(self.rdic, 'BDAT'))
        self.grid.SetCellValue(11, 36, JoinVal(self.rdic, 'BINF'))
        self.grid.SetCellValue(12, 36, JoinVal(self.rdic, 'DPNO'))
        self.grid.SetCellValue(13, 36, JoinVal(self.rdic, 'DPTO'))
        
        self.grid.SetCellValue(9, 36, '_VLS02')
        self.grid.SetCellValue(14, 36, '_VLE02')
        
        self.grid.SetCellValue(14, 0, '_HLS02')
        self.grid.SetCellValue(14, 54, '_HLE02')
        self.grid.SetCellValue(15, 0, '_HLS03')
        self.grid.SetCellValue(15, 54, '_HLE03')
        
        ######################################
        
        self.grid.SetCellValue(15, 1, JoinVal(self.rdic, 'HSNC'))
        self.grid.SetCellValue(16, 4, '_VLS03')
        self.grid.SetCellValue(63, 4, '_VLE03')
        self.grid.SetCellValue(15, 4, JoinVal(self.rdic, 'ITMN'))
        self.grid.SetCellValue(16, 16, '_VLS04')
        self.grid.SetCellValue(63, 16, '_VLE04')
        self.grid.SetCellValue(15, 16, JoinVal(self.rdic, 'IPAC'))
        self.grid.SetCellValue(16, 19, '_VLS05')
        self.grid.SetCellValue(63, 19, '_VLE05')
        self.grid.SetCellValue(15, 19, JoinVal(self.rdic, 'IQTY'))
        self.grid.SetCellValue(16, 24, '_VLS06')
        self.grid.SetCellValue(63, 24, '_VLE06')
        self.grid.SetCellValue(15, 24, JoinVal(self.rdic, 'IBON'))
        self.grid.SetCellValue(16, 27, '_VLS07')
        self.grid.SetCellValue(63, 27, '_VLE07')
        self.grid.SetCellValue(15, 27, JoinVal(self.rdic, 'IRAT'))
        self.grid.SetCellValue(16, 30, '_VLS08')
        self.grid.SetCellValue(63, 30, '_VLE08')
        self.grid.SetCellValue(15, 30, JoinVal(self.rdic, 'IRA2'))
        self.grid.SetCellValue(16, 33, '_VLS09')
        self.grid.SetCellValue(63, 33, '_VLE09')
        self.grid.SetCellValue(15, 33, JoinVal(self.rdic, 'IAMT'))
        self.grid.SetCellValue(16, 38, '_VLS10')
        self.grid.SetCellValue(63, 38, '_VLE10')
        self.grid.SetCellValue(15, 38, JoinVal(self.rdic, 'ITAX'))
        self.grid.SetCellValue(16, 41, '_VLS11')
        self.grid.SetCellValue(63, 41, '_VLE11')
        self.grid.SetCellValue(15, 41, JoinVal(self.rdic, 'IDIS'))
        self.grid.SetCellValue(16, 44, '_VLS12')
        self.grid.SetCellValue(63, 44, '_VLE12')
        self.grid.SetCellValue(15, 44, JoinVal(self.rdic, 'IMRP'))
        self.grid.SetCellValue(16, 47, '_VLS13')
        self.grid.SetCellValue(63, 47, '_VLE13')
        self.grid.SetCellValue(15, 47, JoinVal(self.rdic, 'IEXP'))
        self.grid.SetCellValue(16, 50, '_VLS14')
        self.grid.SetCellValue(63, 50, '_VLE14')
        self.grid.SetCellValue(15, 50, JoinVal(self.rdic, 'INET'))
        self.grid.SetCellValue(16, 53, '_VLS15')
        self.grid.SetCellValue(63, 53, '_VLE15')
        self.grid.SetCellValue(15, 53, JoinVal(self.rdic, 'IBAT'))
        #######################################
        
        self.grid.SetCellValue(63, 0, '_HLS04')
        self.grid.SetCellValue(63, 54, '_HLE04')
        self.grid.SetCellValue(64, 0, '_HLS05')
        self.grid.SetCellValue(64, 54, '_HLE05')
        
        self.grid.SetCellValue(64, 1, JoinVal(self.rdic, 'BTXC'))
        self.grid.SetCellValue(64, 4, JoinVal(self.rdic, 'BIGT'))
        self.grid.SetCellValue(64, 8, JoinVal(self.rdic, 'BCGT'))
        self.grid.SetCellValue(64, 13, JoinVal(self.rdic, 'BSGT')) 
        self.grid.SetCellValue(64, 18, JoinVal(self.rdic, 'TAXP'))
        self.grid.SetCellValue(64, 23, JoinVal(self.rdic, 'BDIS'))
        self.grid.SetCellValue(64, 28, JoinVal(self.rdic, 'BAMT'))
        self.grid.SetCellValue(64, 33, JoinVal(self.rdic, 'BTOT'))
        
        self.grid.SetCellValue(64, 38, '_VLS16')
        self.grid.SetCellValue(71, 38, '_VLE16')
        
        self.grid.SetCellValue(64, 39, JoinVal(self.rdic, 'ITMC'))
        self.grid.SetCellValue(64, 43, JoinVal(self.rdic, 'QTTC'))

        self.grid.SetCellValue(65, 39, JoinVal(self.rdic, 'DISS'))
        self.grid.SetCellValue(66, 39, JoinVal(self.rdic, 'AMTS'))
        self.grid.SetCellValue(67, 39, JoinVal(self.rdic, 'OTH1'))
        self.grid.SetCellValue(68, 39, JoinVal(self.rdic, 'OTH2'))
        self.grid.SetCellValue(69, 39, JoinVal(self.rdic, 'OTH3'))
         
        self.grid.SetCellValue(71, 39, JoinVal(self.rdic, 'GTOT'))

        self.grid.SetCellValue(72, 1, JoinVal(self.rdic, 'AWRD'))
        self.grid.SetCellValue(72, 0, '_HLS07')
        self.grid.SetCellValue(72, 54, '_HLE07')
        
        self.grid.SetCellValue(73, 1, JoinVal(self.rdic, 'CMNT'))

        self.grid.SetCellValue(73, 38, JoinVal(self.rdic, 'AUSF'))
        self.grid.SetCellValue(73, 40, JoinVal(self.rdic, 'AUSO'))
        self.grid.SetCellValue(76, 42, JoinVal(self.rdic, 'AUSS'))
        
        self.grid.SetCellValue(73, 1, JoinVal(self.rdic, 'STA1'))
        self.grid.SetCellValue(74, 1, JoinVal(self.rdic, 'STA2'))
        self.grid.SetCellValue(75, 1, JoinVal(self.rdic, 'STA3'))
        #self.grid.SetCellValue(78, 1, JoinVal(self.rdic, 'STA4'))
        
        self.grid.SetCellValue(69, 0, '_HLS08')
        self.grid.SetCellValue(69, 54, '_HLE08')
        self.grid.SetCellValue(71, 0, '_HLS09')
        self.grid.SetCellValue(71, 54, '_HLE09')
        self.grid.SetCellValue(77, 0, '_END0')
        
    def SetDefaultValue_Half(self):
        self.fh1.SetValue(str('13.0'))
        self.fh2.SetValue(str('12.0'))
        self.spbtline.SetValue('10.0')
        #self.grid.SetCellValue(2, 0, '_HLS01')
        #self.grid.SetCellValue(2, 54, '_HLE01')
        
        self.grid.SetCellValue(1, 27, JoinVal(self.rdic, 'THDP'))
        self.grid.SetCellValue(2, 26, JoinVal(self.rdic, 'CSCR'))
        
        self.grid.SetCellValue(3, 1, JoinVal(self.rdic, 'ONDP'))
        self.grid.SetCellValue(4, 1, JoinVal(self.rdic, 'OAD1'))
        self.grid.SetCellValue(4, 12, JoinVal(self.rdic, 'OAD2'))
        self.grid.SetCellValue(5, 17, JoinVal(self.rdic, 'OPHN'))
        self.grid.SetCellValue(5, 1, JoinVal(self.rdic, 'OREG'))
        self.grid.SetCellValue(6, 1, JoinVal(self.rdic, 'OGST'))
        self.grid.SetCellValue(6, 9, JoinVal(self.rdic, 'OEML'))
        
        self.grid.SetCellValue(3, 22, JoinVal(self.rdic, 'PNDP'))
        self.grid.SetCellValue(4, 22, JoinVal(self.rdic, 'PAD1'))
        self.grid.SetCellValue(4, 30, JoinVal(self.rdic, 'PAD2'))
        self.grid.SetCellValue(5, 22, JoinVal(self.rdic, 'PREG'))
        self.grid.SetCellValue(5, 29, JoinVal(self.rdic, 'PPHN'))
        self.grid.SetCellValue(6, 22, JoinVal(self.rdic, 'PGST'))
        self.grid.SetCellValue(6, 29, JoinVal(self.rdic, 'PEML'))
        
        self.grid.SetCellValue(3, 45, JoinVal(self.rdic, 'BILN'))
        self.grid.SetCellValue(4, 45, JoinVal(self.rdic, 'BDAT'))
        self.grid.SetCellValue(5, 45, JoinVal(self.rdic, 'ORDN'))
        self.grid.SetCellValue(6, 45, JoinVal(self.rdic, 'PMOD'))

        self.grid.SetCellValue(6, 0, '_HLS02')
        self.grid.SetCellValue(6, 54, '_HLE02')

        self.grid.SetCellValue(7, 0, '_HLS03')
        self.grid.SetCellValue(7, 54, '_HLE03')
        
        self.grid.SetCellValue(7, 1, JoinVal(self.rdic, 'HSNC'))
        self.grid.SetCellValue(8, 4, '_VLS03')
        self.grid.SetCellValue(26, 4, '_VLE03')
        self.grid.SetCellValue(7, 4, JoinVal(self.rdic, 'ITMN'))
        self.grid.SetCellValue(8, 16, '_VLS04')
        self.grid.SetCellValue(26, 16, '_VLE04')
        self.grid.SetCellValue(7, 16, JoinVal(self.rdic, 'IPAC'))
        self.grid.SetCellValue(8, 19, '_VLS05')
        self.grid.SetCellValue(26, 19, '_VLE05')
        self.grid.SetCellValue(7, 19, JoinVal(self.rdic, 'IQTY'))
        self.grid.SetCellValue(8, 24, '_VLS06')
        self.grid.SetCellValue(26, 24, '_VLE06')
        self.grid.SetCellValue(7, 24, JoinVal(self.rdic, 'IBON'))
        self.grid.SetCellValue(8, 27, '_VLS07')
        self.grid.SetCellValue(26, 27, '_VLE07')
        self.grid.SetCellValue(7, 27, JoinVal(self.rdic, 'IRAT'))
        self.grid.SetCellValue(8, 30, '_VLS08')
        self.grid.SetCellValue(26, 30, '_VLE08')
        self.grid.SetCellValue(7, 30, JoinVal(self.rdic, 'IRA2'))
        self.grid.SetCellValue(8, 33, '_VLS09')
        self.grid.SetCellValue(26, 33, '_VLE09')
        self.grid.SetCellValue(7, 33, JoinVal(self.rdic, 'IAMT'))
        
        self.grid.SetCellValue(8, 38, '_VLS10')
        self.grid.SetCellValue(26, 38, '_VLE10')
        self.grid.SetCellValue(7, 38, JoinVal(self.rdic, 'ITAX'))
        
        self.grid.SetCellValue(8, 41, '_VLS11')
        self.grid.SetCellValue(26, 41, '_VLE11')
        self.grid.SetCellValue(7, 41, JoinVal(self.rdic, 'IDIS'))
        
        self.grid.SetCellValue(8, 44, '_VLS12')
        self.grid.SetCellValue(26, 44, '_VLE12')
        self.grid.SetCellValue(7, 44, JoinVal(self.rdic, 'IMRP'))
        self.grid.SetCellValue(8, 47, '_VLS13')
        self.grid.SetCellValue(26, 47, '_VLE13')
        self.grid.SetCellValue(7, 47, JoinVal(self.rdic, 'IEXP'))
        self.grid.SetCellValue(8, 50, '_VLS14')
        self.grid.SetCellValue(26, 50, '_VLE14')
        self.grid.SetCellValue(7, 50, JoinVal(self.rdic, 'INET'))
        self.grid.SetCellValue(8, 53, '_VLS15')
        self.grid.SetCellValue(26, 53, '_VLE15')
        self.grid.SetCellValue(7, 53, JoinVal(self.rdic, 'IBAT'))
        #######################################
        self.grid.SetCellValue(26, 0, '_HLS04')
        self.grid.SetCellValue(26, 54, '_HLE04')
        self.grid.SetCellValue(27, 0, '_HLS05')
        self.grid.SetCellValue(27, 54, '_HLE05')
        
        self.grid.SetCellValue(27, 1, JoinVal(self.rdic, 'BTXC'))
        self.grid.SetCellValue(27, 4, JoinVal(self.rdic, 'BIGT'))
        self.grid.SetCellValue(27, 8, JoinVal(self.rdic, 'BCGT'))
        self.grid.SetCellValue(27, 13, JoinVal(self.rdic, 'BSGT')) 
        self.grid.SetCellValue(27, 18, JoinVal(self.rdic, 'TAXP'))
        self.grid.SetCellValue(27, 23, JoinVal(self.rdic, 'BDIS'))
        self.grid.SetCellValue(27, 28, JoinVal(self.rdic, 'BAMT'))
        self.grid.SetCellValue(27, 33, JoinVal(self.rdic, 'BTOT'))
        self.grid.SetCellValue(27, 38, '_VLS16')
        self.grid.SetCellValue(34, 38, '_VLE16')
        
        self.grid.SetCellValue(27, 0, '_HLS06')
        self.grid.SetCellValue(27, 54, '_HLE06')

        self.grid.SetCellValue(27, 39, JoinVal(self.rdic, 'ITMC'))
        self.grid.SetCellValue(27, 43, JoinVal(self.rdic, 'QTTC'))

        self.grid.SetCellValue(28, 39, JoinVal(self.rdic, 'DISS'))
        self.grid.SetCellValue(29, 39, JoinVal(self.rdic, 'AMTS'))
        self.grid.SetCellValue(30, 39, JoinVal(self.rdic, 'OTH1'))
        self.grid.SetCellValue(31, 39, JoinVal(self.rdic, 'OTH2'))
        ##self.grid.SetCellValue(32, 47, JoinVal(self.rdic, 'OTH3'))
        
        self.grid.SetCellValue(33, 39, JoinVal(self.rdic, 'GTOT'))

        self.grid.SetCellValue(34, 1, JoinVal(self.rdic, 'AWRD'))
        self.grid.SetCellValue(34, 0, '_HLS07')
        self.grid.SetCellValue(34, 54, '_HLE07')

        #self.grid.SetCellValue(32, 0, '_HLS08')
        #self.grid.SetCellValue(32, 54, '_HLE08')
        
        self.grid.SetCellValue(35, 1, JoinVal(self.rdic, 'CMNT'))

        self.grid.SetCellValue(35, 38, JoinVal(self.rdic, 'AUSF'))
        self.grid.SetCellValue(35, 40, JoinVal(self.rdic, 'AUSO'))
        self.grid.SetCellValue(38, 42, JoinVal(self.rdic, 'AUSS'))
        
        self.grid.SetCellValue(36, 1, JoinVal(self.rdic, 'STA1'))
        self.grid.SetCellValue(37, 1, JoinVal(self.rdic, 'STA2'))
        self.grid.SetCellValue(38, 1, JoinVal(self.rdic, 'STA3'))
        #self.grid.SetCellValue(49, 1, JoinVal(self.rdic, 'STA4'))
        
        #self.grid.SetCellValue(33, 0, '_HLS09')
        #self.grid.SetCellValue(33, 54, '_HLE09')
        
        self.grid.SetCellValue(39, 0, '_END0')

    def SetDefaultSize8cm(self):
        HSL, HEL = 0, 18 ## Horizontal Start Line, Horizontal End Line
        self.fh1.SetValue(str('11.0'))
        self.fh2.SetValue(str('9.0'))
        self.spbtline.SetValue('8.0')
        self.rdic['HSNC'][0] = 'HSN'
        self.rdic['PGST'][0] = ''
        self.rdic['PGST'] = ['','']
        self.rdic['GTOT'] = ['9*2','']
        self.grid.SetCellValue(1, 4, JoinVal(self.rdic, 'THDP'))
        self.grid.SetCellValue(2, 5, JoinVal(self.rdic, 'CSCR'))
        
        self.grid.SetCellValue(3, 1, JoinVal(self.rdic, 'ONDP'))
        self.grid.SetCellValue(4, 1, JoinVal(self.rdic, 'OAD1'))
        self.grid.SetCellValue(4, 8, JoinVal(self.rdic, 'OAD2'))
        self.grid.SetCellValue(5, 1, JoinVal(self.rdic, 'OGST'))
        self.grid.SetCellValue(5, 8, JoinVal(self.rdic, 'OPHN'))
        self.grid.SetCellValue(6, 1, JoinVal(self.rdic, 'OREG'))
        self.grid.SetCellValue(6, HSL, '_HLS01')
        self.grid.SetCellValue(6, HEL, '_HLE01')
        self.grid.SetCellValue(7, 1, JoinVal(self.rdic, 'BILN'))
        self.grid.SetCellValue(7, 8, JoinVal(self.rdic, 'BDAT'))
        self.grid.SetCellValue(8, HSL, '_HLS02')
        self.grid.SetCellValue(8, HEL, '_HLE02')
        ##self.grid.SetCellValue(11, 1, JoinVal(self.rdic, 'PNDP'))
        self.grid.SetCellValue(9, 1, JoinVal(self.rdic, 'PAD1'))## PAD1 = Customer Name
        self.grid.SetCellValue(9, 9, JoinVal(self.rdic, 'PGST'))
        #self.grid.SetCellValue(9, 8, JoinVal(self.rdic, 'PREG'))## Phone Number
        #self.grid.SetCellValue(10, HSL, '_HLS02')
        #self.grid.SetCellValue(10, HEL, '_HLE02')

        self.grid.SetCellValue(10, HSL, '_HLS03')
        self.grid.SetCellValue(10, HEL, '_HLE03')

        self.grid.SetCellValue(11, HSL, '_HLS04')
        self.grid.SetCellValue(11, HEL, '_HLE04')
        
        self.grid.SetCellValue(11, 1, JoinVal(self.rdic, 'HSNC'))
        self.grid.SetCellValue(12, 3, '_VLS03')
        self.grid.SetCellValue(22, 3, '_VLE03')
        self.grid.SetCellValue(11, 3, JoinVal(self.rdic, 'ITMN'))
        
        self.grid.SetCellValue(12, 11, '_VLS04')
        self.grid.SetCellValue(22, 11, '_VLE04')
        self.grid.SetCellValue(11, 11, JoinVal(self.rdic, 'IQTY'))
        
        self.grid.SetCellValue(12, 13, '_VLS06')
        self.grid.SetCellValue(22, 13, '_VLE06')
        self.grid.SetCellValue(11, 13, JoinVal(self.rdic, 'IAMT'))
        
        self.grid.SetCellValue(12, 17, '_VLS07')
        self.grid.SetCellValue(22, 17, '_VLE07')
        self.grid.SetCellValue(11, 17, JoinVal(self.rdic, 'IRAT'))
        #self.grid.SetCellValue(14, 27, '_VLS08')
        #self.grid.SetCellValue(26, 27, '_VLE08')
       
        #######################################
        self.grid.SetCellValue(22, HSL, '_HLS05')
        self.grid.SetCellValue(22, HEL, '_HLE05')
        self.grid.SetCellValue(23, HSL, '_HLS06')
        self.grid.SetCellValue(23, HEL, '_HLE06')
        
        self.grid.SetCellValue(23, 1, JoinVal(self.rdic, 'BTXC'))
        self.grid.SetCellValue(23, 51, JoinVal(self.rdic, 'BIGT'))## Ignored Send Outside Prit Area
        self.grid.SetCellValue(23, 52, JoinVal(self.rdic, 'TAXP'))## Ignored Send Outside Prit Area
        self.grid.SetCellValue(23, 53, JoinVal(self.rdic, 'BDIS'))## Ignored Send Outside Prit Area
        
        self.grid.SetCellValue(23, 4, JoinVal(self.rdic, 'BCGT'))
        self.grid.SetCellValue(23, 8, JoinVal(self.rdic, 'BSGT')) 
        self.grid.SetCellValue(23, 12, JoinVal(self.rdic, 'BAMT'))
        self.grid.SetCellValue(23, 16, JoinVal(self.rdic, 'BTOT'))
        
        self.grid.SetCellValue(23, HSL, '_HLS07')
        self.grid.SetCellValue(23, HEL, '_HLE07')

        self.grid.SetCellValue(29, HSL, '_HLS08')
        self.grid.SetCellValue(29, HEL, '_HLE08')
        
        self.grid.SetCellValue(30, 1, JoinVal(self.rdic, 'DISS'))
        self.grid.SetCellValue(30, 10, JoinVal(self.rdic, 'AMTS'))

        self.grid.SetCellValue(31, 1, JoinVal(self.rdic, 'ITMC'))
        self.grid.SetCellValue(32, 1, JoinVal(self.rdic, 'QTTC'))

        self.grid.SetCellValue(32, 9, JoinVal(self.rdic, 'GTOT'))

        self.grid.SetCellValue(33, 1, JoinVal(self.rdic, 'AWRD'))
        self.grid.SetCellValue(33, HSL, '_HLS09')
        self.grid.SetCellValue(33, HEL, '_HLE09')

        self.grid.SetCellValue(34, 1, JoinVal(self.rdic, 'CMNT'))

        self.grid.SetCellValue(35, 1, JoinVal(self.rdic, 'STA1'))
        self.grid.SetCellValue(36, 1, JoinVal(self.rdic, 'STA2'))
        self.grid.SetCellValue(37, 1, JoinVal(self.rdic, 'STA3'))

        self.grid.SetCellValue(38, 1, JoinVal(self.rdic, 'AUSF'))
        self.grid.SetCellValue(38, 7, JoinVal(self.rdic, 'AUSO'))
        self.grid.SetCellValue(39, 9, JoinVal(self.rdic, 'AUSS'))
        
        self.grid.SetCellValue(39, HSL, '_HLS10')
        self.grid.SetCellValue(39, HEL, '_HLE10')
        
        self.grid.SetCellValue(40, 0, '_END0')
        
        self.fh1.ChangeValue('11') ## Head Font Size, Owner Name 
        self.fh2.ChangeValue('9') ## Head2 Font Size, Customer Name 
        self.fh3.ChangeValue('9') ##BillNo,BillDate,DiscountAmount,TotalAmount,OtherSum
        self.fitm.ChangeValue('7') ##Item Font Size Heading, Authorized Signatory
        self.fh4.ChangeValue('7')##Font Size Used in \nCash-Credit, OwnerDetails, ItemCount, QuantityCount,
                                  ##Amount in Words, Comments, Authorised Owner
        self.fh5.ChangeValue('7') ##Font Size Used in Disclaimers 
        self.fh6.ChangeValue('7') ##Font Size Used in GST TAX Summary 
        self.fn2pt.ChangeValue('7') ## Font Size Used in Party Details 
        self.fns.ChangeValue('7') ## Font Size Used in Top Greating Message Page Top
        self.fnsn.ChangeValue('Times-Roman')
        numrows = 41
        self.getrowline.ChangeValue(str(numrows))
        self.httx.ChangeValue(str(numrows*10))
        self.pwtx.ChangeValue('225')
        
    def Restlinedic(self, mcol):
        #self.hld = {''}
        self.hld = {'LPRT':[self.lanpot,self.lanpot,self.rsendline.GetValue()],}
        self.vld = {}
        for i in range(1, mcol):
            self.hld['_HLS%02d'%i] = ['','', '']
            self.hld['_HLE%02d'%i] = ['','', '']
            self.vld['_VLS%02d'%i] = ['','', '']
            self.vld['_VLE%02d'%i] = ['','', '']
        self.vld['_END0'] = ['','', '']
        
    def Restrdic(self):
        try:
            pw = float(self.pwtx.GetValue())
        except:
            pw = '600'
        try:
            ph = float(self.httx.GetValue())
        except:
            ph = '820'
        
        self.rdic = {'LOGO':['',''], 'THDP':['',''], 'CSCR':['9*1', ''], 'ONDP':['','me'],
        'OAD1':['',''], 'OAD2':['', ''], 'OPHN':['',''],
        'OEML':['',''], 'OREG':['',''], 'OGST':['',''],
        'PNDP':['', ''], 'PAD1':['',''], 'PAD2':['',''], 'PPHN':['',''],
        'PEML':['',''], 'PREG':['',''], 'PGST':['GSTN:',''],'BILN':['Bill No:',''],
        'BDAT':['Bill Date:',''],'DITM':['Deliver ITems:',''], 'BINF':['',''], 'AWRD':['Rs.',''],
        'ORDN':['Order No.:',''],'DDAT':['',''],'PMOD':['Pay.Mode:',''], 
        'DPNO':['Dispatch No.:',''],'DPTO':['Dispatched To:',''],'HLIN':['','horizontal line'], 'VLIN':['','vertical line'], 
        'HSNC':['HSNC',''],'ITMN':['ITEM NAME',''],'IPAC':['Pack',''], 'IQTY':['Qty',''], 'IBON':['Bonus',''], 'IRAT':['Rate',''], 'IRA2':['Rate2',''],
        'IAMT':['Amount',''], 'IMRP':['MRP',''],'ITAX':['Tax%',''],'IDIS':['Dis%',''],'IEXP':['Exp.',''],
        'INET':['Net',''],'IBAT':['Batch',''],'PAGESIZE':(pw, ph), 'BTXC':['GST%', ''],
        'BCGT':['CGST', ''],  'BIGT':['IGST',''], 'BSGT':['SGST',''], 'TAXP':['Tax Amt.',''], 'BDIS':['Discount',''], 'BAMT':['Amount',''],
        'BTOT':['Total',''], 'ITMC':['Items:',''], 'QTTC':['Total Qty:',''], 'STA1':['',''],'STA2':['',''],'STA3':['',''], 
        'STA4':['',''],'AUSS':['Authorized Signatory',''], 'AUTS':['',''], 'CMNT':['MSG : ',''], 'ADTS':['Add Taxes',''], 
        'OTH1':['',''],'OTH2':['',''],'OTH3':['',''],'DISS':['Discount : ',''],'AMTS':['Amount : ',''],
        'GTOT':['18*2',''], 'AUSO':['',''], 'AUSF':['For:','']}

    def SaveDisable(self):
        self.save.Disable()
        self.compsave.Disable()
        
    def OnCompSave(self, event):
        itms, rdic, hld, vld, fnt, fnm, spbtline, gtothead = self.DicInfo()
        infolst = [hld, vld, spbtline, gtothead]
        ####stw = PRINT_DISCRIPTION()
        comp = True
        self.PDC.FILE_WRITE(rdic, hld, vld, fnt, fnm, spbtline, gtothead, comp)
        self.SaveDisable()
        stw.READFILE()
        self.RelodUpdated(self.PDC)
        self.PDC = PRINT_DISCRIPTION() 
        event.Skip()
        
    def OnSave(self, event):
        itms, rdic, hld, vld, fnt, fnm, spbtline, gtothead = self.DicInfo()
        infolst = [hld, vld, spbtline, gtothead]
        ###stw = PRINT_DISCRIPTION()
        comp = False
        self.PDC.FILE_WRITE(rdic, hld, vld, fnt, fnm, spbtline, gtothead, comp)
        self.SaveDisable()
        self.RelodUpdated(self.PDC)
        self.PDC = PRINT_DISCRIPTION() 
        event.Skip()

    def OnSaveindpg4(self, event):
        self.SaveDisable()
        self.saveindpg4.Disable()
        self.load.Enable()
        self.compload.Enable()
        try:
            itms, rdic, hld, vld, fnt, fnm, spbtline, gtothead = self.DicInfo()
            infolst = [hld, vld, spbtline, gtothead]
            comp = False
            self.PDC.FILE_WRITE(rdic, hld, vld, fnt, fnm, spbtline, gtothead, comp, getfile=self.landpgpath4)
            self.SaveDisable()
            self.RelodUpdated(self.PDC)
            self.parent.status.SetStatusText('Succefully Save [ %s ] File'%str(self.landpgpath4))
        except Exception as err:
            self.parent.status.SetStatusText('Some Error Stop Save >>> [%s]'%str(err))
        self.PDC = PRINT_DISCRIPTION() 
        event.Skip()
        
    def OnLoadindpg4(self, event):
        self.SaveDisable()
        self.saveindpg4.Enable()
        try:
            self.loadmsg = True
            self.ClearGrid()
            rdic, hld, vld, fnt, fnm, spbtline, gtothead = self.LoadFile(self.landpgpath4)
            self.LoadAll(rdic, hld, vld, fnt, fnm, spbtline, gtothead)
            self.parent.status.SetStatusText('Succefully Loaded [ %s ] File'%str(self.landpgpath4))
        except Exception as err:
            self.parent.status.SetStatusText('Some Error Stop Loading >>> [%s]'%str(err))
        self.save.Disable()
        self.load.Disable()
        self.compsave.Disable()
        self.compload.Disable()
        event.Skip()
        
    def OnCreate_indpg4(self, event):
        actlab = self.Create_indpg4.GetLabel()
        self.Create_indpg4.SetLabel('Wait ...!!')
        stfile = rmss_config.my_icon(self.PGPATH3)
        self.landpgpath4 = stfile
        if not os.path.exists(rmss_config.my_icon(stfile)):
            configg = ConfigParser.ConfigParser()
            try:
                with open(stfile, 'w') as f:
                    configg.write(f)
            except IOError as e:
                dirpath = os.path.dirname(self.PGPATH3)
                os.makedirs(dirpath)
                self.Create_indpg4.SetLabel('Press Again')
                event.Skip()
                return

        self.LandScapeA4()
        itms, rdic, hld, vld, fnt, fnm, spbtline, gtothead = self.DicInfo()
        infolst = [hld, vld, spbtline, gtothead]
        comp = False
        self.PDC.FILE_WRITE(rdic, hld, vld, fnt, fnm, spbtline, gtothead, comp, getfile=stfile)
        self.SaveDisable() 
        self.load.Disable()
        self.compload.Disable()
        self.Create_indpg4.SetLabel(actlab)
        self.saveindpg4.Disable()
        self.loadindpg4.Enable()
        event.Skip()
        
    def RelodUpdated(self, stw):
        prinfodict = stw.READFILE()
        maxline = 17
        
        if prinfodict:
            if prinfodict[2]['_END0'][1] in range(380, 450): 
                #### DEFAULT pagebottm_margin >>> pagebottm_margin = 420 range(380, 450) for half A4_size_page;
                #### Full A4_size_page has pagebottm_margin = 10
                
                if self.resource_dic['iteminfo']:
                    maxline = 10 ###12 line will default with item discription below itemname
                else:
                    maxline = 17 ###17 line will default
            else:
                if self.resource_dic['iteminfo']:
                    maxline = 22 ###24 line will default with item discription below itemname
                else:
                    maxline = 35 ###17 line will default
        self.resource_dic['printmeth'][1]=prinfodict
        
    def LoadFile(self, getfile=None):
        ###stw = PRINT_DISCRIPTION()
        rdic, hld, vld, fnt, fnm, spbtline, gtothead = self.PDC.READFILE(getfile)
        self.spbtline.SetValue(str(spbtline))
        self.gth.SetValue(str(gtothead))
        return rdic, hld, vld, fnt, fnm, spbtline, gtothead
    
    def LoadCompFile(self):
        stw = PRINT_DISCRIPTION()
        rdic, hld, vld, fnt, fnm, spbtline, gtothead = stw.READFILE_COMP()
        self.spbtline.SetValue(str(spbtline))
        self.gth.SetValue(str(gtothead))
        return rdic, hld, vld, fnt, fnm, spbtline, gtothead
    
    def LoadAll(self, rdic, hld, vld, fnt, fnm, spbtline, gtothead):
        self.fns.SetValue(str(fnt.get('fs1')))
        self.fh4.SetValue(str(fnt.get('fh4')))
        self.fh5.SetValue(str(fnt.get('fh5')))
        self.fh6.SetValue(str(fnt.get('fh6')))
        self.fitm.SetValue(str(fnt.get('fit')))
        self.fitmh.SetValue(str(fnt.get('fitmh')))
        self.fh3.SetValue(str(fnt.get('fh3')))
        self.fh2.SetValue(str(fnt.get('fh2')))
        self.fh1.SetValue(str(fnt.get('fh1')))
        self.fn2pt.SetValue(str(fnt.get('f2pt')))
        #######################################
        self.fnsn.SetValue(str(fnm.get('fs1n')))
        self.fh4n.SetValue(str(fnm.get('fh4n')))
        self.fh5n.SetValue(str(fnm.get('fh5n')))
        self.fh6n.SetValue(str(fnm.get('fh6n')))
        self.fitmn.SetValue(str(fnm.get('fitn')))
        self.fitmnh.SetValue(str(fnm.get('fitnh')))
        self.fh3n.SetValue(str(fnm.get('fh3n')))
        self.fh2n.SetValue(str(fnm.get('fh2n')))
        self.fh1n.SetValue(str(fnm.get('fh1n')))
        self.fn2ptn.SetValue(str(fnm.get('f2ptn')))
        #######################################
        self.gth.SetValue(str(gtothead.strip()))
        self.spbtline.SetValue(str(spbtline))
       
        rdlst = []
        hllst = []
        vllst = []
        for k, v in rdic.iteritems():
            if '__NAME__' not in k:
                try:
                    rdlst.append([k, v[0][1][0], v[0][1][1], v[0][0]])
                except (IndexError, TypeError):
                    pass
        self.pwtx.ChangeValue(str(rdic['PAGESIZE'][0]))
        self.httx.ChangeValue(str(rdic['PAGESIZE'][1]))
        for k, v in hld.iteritems():
            #self.grid.SetCellValue(36, 39, JoinVal(self.rdic, 'ITMC'))
            if '__NAME__' not in k:
                if v[0] != "":
                    try:
                        hllst.append([k, v[0], v[1]])
                        #print k, v[0], v[1]
                    except (IndexError, TypeError):
                        pass
        try: 
            self.lanpot = hld['LPRT'][0] ## Landscape, 1 or Potrait, 2
            self.slider.SetValue(hld['LPRT'][0])
            self.rsendline.SetValue(str(hld['LPRT'][2]))
        except :
            self.lanpot = 2
            self.slider.SetValue(self.lanpot)
            self.rsendline.SetValue('575')
        for k, v in vld.iteritems():
            #self.grid.SetCellValue(36, 39, JoinVal(self.rdic, 'ITMC'))
            if '__NAME__' not in k:
                if v[0] != "":
                    try:
                        vllst.append([k, v[0], v[1]])
                    except (IndexError, TypeError):
                        pass
        col = self.grid.GetNumberCols()
        try:
            row = int(rdic['PAGESIZE'][1]/10)
        except:
            row = self.grid.GetNumberRows()
        self.getrowline.ChangeValue(str(row))
        for r in (range(row)):
            #print r, rr
            try:
                #self.grid.SetCellValue(rr, 39, JoinVal(self.rdic, 'ITMC'))
                ro = row-(rdlst[r][2]/10.0)
                #co = rdlst[r][1]/10.0
                co = sum([(rdlst[r][1]/10.0), -1])
                #print rdlst[r][0], rdlst[r][1]/10.0, row-(rdlst[r][2]/10.0), rdlst[r][3]
                self.grid.SetCellValue(ro, co, ':'.join([rdlst[r][0], rdlst[r][3]]))
            except (IndexError, TypeError):
                pass
        modrow = row-1
        for r in (range(row)):
            hndrow = row
            try:
                #self.grid.SetCellValue(rr, 39, JoinVal(self.rdic, 'ITMC'))
                ro = hndrow-(hllst[r][2]/10.0)
                co = sum([(hllst[r][1]/10.0), -1])
                self.grid.SetCellValue(ro, co, hllst[r][0].upper())
            except (IndexError, TypeError):
                pass
        for r in (range(row)):
            try:
                #self.grid.SetCellValue(rr, 39, JoinVal(self.rdic, 'ITMC'))
                ro = row-(vllst[r][2]/10.0)
                co = sum([(vllst[r][1]/10.0), -1])
                #co = vllst[r][1]/10.0
                #print rdlst[r][0], rdlst[r][1]/10.0, row-(rdlst[r][2]/10.0), rdlst[r][3]
                if '_END' in (vllst[r][0].upper()):
                    self.grid.SetCellValue(ro, co, vllst[r][0].upper())
                else:
                    self.grid.SetCellValue(ro, co, vllst[r][0].upper())
            except (IndexError, TypeError):
                pass
        self.rdic.clear()
        self.hld.clear()
        self.vld.clear()
        self.rdic = rdic
        self.hld = hld
        self.vld = vld
        
    def OnLoad(self, event):
        try:
            self.loadmsg = True
            self.ClearGrid()
            rdic, hld, vld, fnt, fnm, spbtline, gtothead = self.LoadFile()
            self.LoadAll(rdic, hld, vld, fnt, fnm, spbtline, gtothead)
            self.save.Enable()
            self.compsave.Disable()
            self.loadindpg4.Disable()
            self.saveindpg4.Disable()
        except ValueError:
            pass
        event.Skip()
        
    def OnCompLoad(self, event):
        try:
            self.loadmsg = True
            self.ClearGrid()
            rdic, hld, vld, fnt, fnm, spbtline, gtothead = self.LoadCompFile()
            self.LoadAll(rdic, hld, vld, fnt, fnm, spbtline, gtothead)
            self.save.Disable()
            self.compsave.Enable()
            self.loadindpg4.Disable()
            self.saveindpg4.Disable()
        except ValueError:
            pass
        event.Skip()
        
    def FontLoad(self):
        try:
            fnt = {'fs1':float(self.fns.GetValue()), 'fh4':float(self.fh4.GetValue()), 'fit':float(self.fitm.GetValue()), 'fitmh':float(self.fitmh.GetValue()),
                'f2pt':float(self.fn2pt.GetValue()), 'fh3':float(self.fh3.GetValue()), 'fh2':float(self.fh2.GetValue()), 'fh1':float(self.fh1.GetValue()),
                   'fh5':float(self.fh5.GetValue()), 'fh6':float(self.fh6.GetValue())}
            fnm = {'fs1n':self.fnsn.GetValue(), 'fh4n':self.fh4n.GetValue(), 'fitn':self.fitmn.GetValue(), 'fitnh':self.fitmnh.GetValue(),
                'f2ptn':self.fn2ptn.GetValue(),'fh3n':self.fh3n.GetValue(), 'fh2n':self.fh2n.GetValue(), 'fh1n':self.fh1n.GetValue(),
                   'fh5n':self.fh5n.GetValue(), 'fh6n':self.fh6n.GetValue()}
        except:
            fnt = {'fs1':7,'fh4':8,'fit':9,'fitmh':9,'f2pt':9,'fh3':9,'fh2':15,'fh1':20,'fh5':8,'fh6':8,}
            fnm = {'fs1n':self.fnptnm,'fh4n':self.fh4nm,'fitn':self.fitmnm,'fitnh':self.fnptnm,'f2ptn':self.fnptnm,
                   'fh3n':self.fh3nm,'fh2n':self.fh2nm,'fh1n':self.fh1nm,'fh5n':self.fh4nm,'fh6n':self.fh4nm,}
        return fnt, fnm
    
    def DicInfo(self):
        #row = self.grid.GetNumberRows()
        col = self.grid.GetNumberCols()
        row = int(self.getrowline.GetValue())
        mm = 10
        self.Restrdic()
        self.Restlinedic(self.mcol)
        for r, rr in enumerate(range(row, 0, -1)):
            for c in range(col):
                if self.grid.GetCellValue(r, c) != "":
                    try:
                        if self.grid.GetCellValue(r, c).strip()[5:]== "":
                            self.rdic[self.grid.GetCellValue(r, c).upper()[:4]][0] = [self.rdic[self.grid.GetCellValue(r, c).upper()[:4]][0], [sum([c,1])*mm, rr*mm]]
                        else:
                            self.rdic[self.grid.GetCellValue(r, c).upper()[:4]][0] = [self.grid.GetCellValue(r, c).strip()[5:], [sum([c,1])*mm, rr*mm]]
                    except KeyError:
                        pass
                #print self.hld.get(self.grid.GetCellValue(r, c).upper())
                if self.hld.get(self.grid.GetCellValue(r, c).upper()) != None:
                    lgv = self.grid.GetCellValue(r, c).upper()
                    try:
                        gint = int(lgv[4:6])
                        self.hld[lgv][0] = sum([c,1])*mm
                        self.hld[lgv][1] = rr*mm
                    except :
                        pass
                if self.vld.get(self.grid.GetCellValue(r, c).upper()) != None:
                    vgv = self.grid.GetCellValue(r, c).upper()
                    try:
                        gint = int(vgv[4:6])
                        self.vld[vgv][0] = sum([c,1])*mm
                        self.vld[vgv][1] = rr*mm 
                    except :
                        pass
        
        fnt, fnm = self.FontLoad()
        try:
            spbtline = float(self.spbtline.GetValue())
        except ValueError:
            spbtline = 2
        ###itms = billinfo['prodtable']
        itms = self.itmlst1
        rdic = self.rdic
        hld = self.hld
        vld = self.vld
        gtothead = self.gth.GetValue().strip()
        #print rdic
        return itms, rdic, hld, vld, fnt, fnm, spbtline, gtothead

    def OnClose(self, event):
        self.parent.Close()
        self.parent.Destroy()
        event.Skip()
        
    def CompChk(self):
        ##compval, stockist_det = OWNER_STOKIST_FOR_With_Other()
        compval = self.resource_dic['owner'][11][0]
        ###stockist_det = self.resource_dic['owner'][11][1]
        try:
            com_p = str(compval).split('/')
            comp = com_p[1]
            if comp.upper() == 'FALSE':
                comp = False
                rdic, hld, vld, fnt, fnm, spbtline, gtothead = self.resource_dic['printmeth'][1]
            else:
                comp = True
                rdic, hld, vld, fnt, fnm, spbtline, gtothead = PRINT_DISCRIPTION().READFILE_COMP()
            thdp = com_p[0]
        except IndexError:
            comp, thdp = False, ""
            rdic, hld, vld, fnt, fnm, spbtline, gtothead = self.resource_dic['printmeth'][1]
        BMAIN = True
        return comp, thdp, None, BMAIN, rdic, hld, vld, fnt, fnm, spbtline, gtothead
    
    def OnPrint(self, event):
        if self.loadmsg == False :
            itms, rdic, hld, vld, fnt, fnm, spbtline, gtothead = self.DicInfo()
            #rdic = self.LoadRdicPrint(rdic)
            compval, thdp, suplystat, BMAIN, rdic, hld, vld, fnt, fnm, spbtline, gtothead = self.CompChk()
        else:
            #self.Restrdic()
            #self.Restlinedic(self.mcol)
            itms = self.itmlst1
            rdic = self.rdic
            hld = self.hld
            vld = self.vld
            fnt, fnm = self.FontLoad()
            spbtline = float(self.spbtline.GetValue())
            gtothead = self.gth.GetValue()
            rdic = self.LoadRdicPrint(rdic)
        
        igstbool = False
        compval = False  ### Composition Value Check
        pdf_open = True
        BMAIN = False
        try:
            MRECEIPT_PDF(compval, self.resource_dic, 'BillPrint_Test.pdf',itms, BMAIN, rdic, hld, vld, fnt, fnm, spbtline, gtothead, igstbool, pdf_open, parent=self )
        #except Exception as err:
        except KeyboardInterrupt as err:
            wx.MessageBox('Close BillPrint_Test.pdf And Open Again,\n'
                '%s'%err, 'BillPrint_Test.pdf is Already Open',wx.ICON_EXCLAMATION)
        event.Skip()
        
    def LoadRdicPrint(self, rdic):
        rdic['THDP'][1] = 'BILL HEADING'
        rdic['CSCR'][1] = ''.join(['TAX INVOICE', '-', 'CASH'])
        rdic['PNDP'][1] = 'SAMPLE PARTY CUSTOMER NAME'
        rdic['PAD1'][1] = 'SAMPLE PARTY ADDRESS 1'.title()
        rdic['PAD2'][1] = 'SAMPLE PARTY ADDRESS 2'.title()
        rdic['PMOD'][1] = self.resource_dic['bankinfo']['name']
        rdic['BINF'][1] = ', '.join([self.resource_dic['bankinfo']['ifsc'], self.resource_dic['bankinfo']['ac']])
        rdic['PPHN'][1] = ''
        rdic['PREG'][1] = 'CUSTOMER REGISTRATION NO:'
        rdic['DITM'][1] = ''
        rdic['DPNO'][1] = 'Dated'
        rdic['PGST'][1] = ', PIN:'.join(['CUSTOMER GSTN ', ''])
        rdic['PEML'][1] = 'customeremail@xyz.com'
        rdic['ORDN'][1] = "Order No"
        rdic['DPTO'][1] = 'Destination'
        rdic['CMNT'][1] = 'Comment For Customer'
        return rdic
    
class PageSetup(wx.Frame):
    instance = None
    init = 0
    def __new__(self, parent, getdsp_size, resource_dic):
        if self.instance is None:
            self.instance = wx.Frame.__new__(self)
        elif isinstance(self.instance, wx._core._wxPyDeadObject):
            self.instance = wx.Frame.__new__(self)
        return self.instance

    #----------------------------------------------------------------------
    def __init__(self, parent, getdsp_size, resource_dic):
        #Constructor
        if self.init:
            self.SetFocus()
            self.Maximize(True)
            return
        self.init = 1
        wx.Frame.__init__(self, parent ,size=(1000,700),pos=(0, 30), style=rmss_config.frame_style)
        favicon = wx.Icon(rmss_config.my_icon('my_icon.ico'), wx.BITMAP_TYPE_ICO, 16, 16)
        wx.Frame.SetIcon(self, favicon)
        self.Maximize(True)
        self.status=self.CreateStatusBar()
        self.panel = PageSetup1(self, getdsp_size, resource_dic)
        self.SetTitle(("BILL PRINT PAGE SETUP, PDF FORMAT"))
        self.Show()
        self.status.SetStatusText(" BILL PRINT PAGE SETUP, PDF FORMAT")
        #self.Bind(wx.EVT_MAXIMIZE, self.OnLayoutNeeded)
        #self.Bind(wx.EVT_CLOSE, self.OnClose)
    def OnLayoutNeeded(self, evt):
        pass
    def OnClose(self, evt):
        pass
        #evt.Skip()
#"""
if __name__ == "__main__":
    app =wx.App(False)
    getdsp_size = 0
    owner = ['RMS DEMO PARTY', 'RMS DEMO ADDRESS1', 'RMS DEMO ADDRESS1', '9999988891',
             'DL/Reg.No', 'GSTN: 1234566777','Disclamer 1', 'Disclamer 2', 'Disclamer 3 ',
             'Extra Line', [60, u'Y', u'Y', u'Y', u'4', 1], ['/True', u'Company Name Display']]
    resource_dic = {'iteminfo':False, 'owner':owner,'oemail':'  ',
                    'onm_on_esti_bill':'TRUE', 'printmeth':[None, None, None],
                    'bankinfo':{'name':'TEST BANK','ifsc':'ABCD123456','ac':'0000123456789'},}
    frame = PageSetup(None, getdsp_size, resource_dic)
    ##frame.Show()
    app.MainLoop()
#"""
