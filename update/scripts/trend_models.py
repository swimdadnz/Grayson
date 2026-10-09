"""Trend-model utilities for the Grayson swim site.

The asymptote is descriptive, not a career forecast. Transform origins are fixed from the
training data so cross-validation predictions are coherent.
"""
from __future__ import annotations
import numpy as np
from scipy.optimize import curve_fit
from scipy.interpolate import UnivariateSpline
MODEL_COMPLEXITY={"linear":2,"quadratic":3,"logarithmic":2,"asymptotic_exponential":3,"stretched_exponential":4,"spline":5}
def fit_model(name,x,y):
    x=np.asarray(x,float); y=np.asarray(y,float); origin=float(np.min(x)); z=x-origin
    if name=="linear":
        co=np.polyfit(z,y,1); return lambda q:np.polyval(co,np.asarray(q,float)-origin),co
    if name=="quadratic":
        co=np.polyfit(z,y,2); return lambda q:np.polyval(co,np.asarray(q,float)-origin),co
    if name=="logarithmic":
        co=np.polyfit(np.log(z+0.25),y,1); return lambda q:np.polyval(co,np.log(np.asarray(q,float)-origin+0.25)),co
    if name=="asymptotic_exponential":
        fn=lambda zz,floor,b,k:floor+b*np.exp(-k*zz)
        p,_=curve_fit(fn,z,y,p0=[max(0,min(y)*.94),max(.1,max(y)-min(y)),.7],bounds=([max(0,min(y)*.55),0,1e-6],[min(y)*1.05,max(y)*3,20]),maxfev=50000)
        return lambda q:fn(np.asarray(q,float)-origin,*p),p
    if name=="stretched_exponential":
        fn=lambda zz,floor,b,k,p:floor+b*np.exp(-np.power(np.maximum(k*zz,0),p))
        p,_=curve_fit(fn,z,y,p0=[max(0,min(y)*.94),max(.1,max(y)-min(y)),.7,1],bounds=([max(0,min(y)*.55),0,1e-6,.2],[min(y)*1.05,max(y)*3,20,4]),maxfev=100000)
        return lambda q:fn(np.asarray(q,float)-origin,*p),p
    if name=="spline":
        o=np.argsort(x); xx=x[o]; yy=y[o]; s=max(len(xx)*np.var(yy)*.35,1e-6); spl=UnivariateSpline(xx,yy,s=s,k=min(3,len(xx)-1)); return lambda q:spl(np.asarray(q,float)),None
    raise ValueError(name)
