
from goto import label

for i in range(1, 4):

    for j in range(1, 4):

        for k in range(1, 4):

            print(j,i*k,k)

            if k == 3:

                goto .end

label .end

print("did a break from a nested for loop")