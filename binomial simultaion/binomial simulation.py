from scipy.stats import binom
from numpy.random import binomial
from numpy import mean
from numpy import linspace
import matplotlib.pyplot as plt
print("For a k out of n system:\n")

k=int(input("Enter a value for k:\n"))
n=int(input("Enter a value for n:\n"))
p=float(input("Enter a value for p:\n"))
round_dp=6

#Analytical probability mass function for binomial distribution
a_res=binom.pmf(k,n,p)
print("The probability of EXACTLY " + str(k)+ " successful trials out of " + str(n) + " trials is: " + str(round(a_res,round_dp)))
print("----------------------------")
a_res=0
a_res_sum=0
for i in range(k,n+1):
    a_res = binom.pmf(i,n,p)
    print("The probability of EXACTLY " + str(i)+ " successful trials out of " + str(n) + " trials is: " + str(round(a_res,round_dp)))
    a_res_sum=a_res_sum+a_res
print("----------------------------")
print("The probability of >=" + str(k)+ " successful trials out of " + str(n) + " trials is: " + str(round(a_res_sum,round_dp)))

#Simulation
sim=input("\n\n\nWould you like to run a simulation for the calculations above? (y/n)\n")
if sim=='y':
    results_mean=[]
    sims=input("Enter the number of simulations per runs, seperated by commas, e.g. 500,1000,20000:\n")
    repeats=int(input("Enter the number of repeats:\n"))
    counter=list(map(int,sims.split(",")))
    for a in counter:
        results=[]
        print("\n")    
        for r in range(repeats):
            res_sum=0
            for i in range(k,n+1):
                res = sum(binomial(n,p,a)==i)/a
                res_sum=res_sum+res
            results.append(res_sum)
        mean_res=mean(results)
        results_mean.append(mean_res)
        print("mean for",str(int(a)),"simulations:")
        print("The probability of >=" + str(k)+ " successful trials out of " + str(n) + " trials is: " + str(round(mean_res,round_dp)))
        xs=linspace(a,a,repeats)
        plt.scatter(xs,results,s=5,c='r')
    plt.scatter(counter,results_mean,s=10,c='b')
    plt.title("K out of N ("+str(k)+"/"+str(n)+")\n Probability of success (Reliability) = "+str(p))
    plt.xlabel("Simulations per point")
    #plt.ylabel("Probability of "+str(k)+" successes (working units) out of "+str(n)+" units")
    plt.ylabel("System Reliability")
    plt.plot(counter,results_mean,linewidth=0.5,c='b')
    plt.plot([min(counter),max(counter)],[a_res_sum,a_res_sum],linewidth=0.5,c='r',linestyle='-')
    plt.grid()
    plt.show()
else:
    exit
input()

#250,500,750,1000,1250,1500,1750,2000,2250,2500,2750,3000,3250,3500,3750,4000,4250,4500,4750,5000,6500,7000,7500,8000,8500,9000,9500,10000,11000,12000,13000,14000,15000,16000,17000,18000,19000,20000,22000,24000,26000,28000,30000,35000,40000,45000,50000,60000,70000,80000,90000,100000
#note that the result may not match exactly the true solution due to the iterative nature of the solution, therefore very high probability results (close to 1) may be incorrect and be calculated as >1


'''we may want to add box plots to this to get some kind of percentile, OR somehow implement confidence bounds?'''
