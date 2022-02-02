# import arrow
# import datetime

# # testTime = arrow.get(datetime.datetime(2013, 5, 5), 'US/Pacific')
# start = arrow.now('US/Central').replace(minute=0, second=0).to('UTC')

# print(start)

d1 = {'a':123,'b':234}

d2 = {'b':5664,'c':45}

alldict = [d1.keys(),d2.keys()]

print(alldict)

intersectionSet = alldict[0]
for i in alldict:
	intersectionSet = intersectionSet & i
	print(intersectionSet)

# allkey = alldict[0].intersection(*alldict)

print(intersectionSet)