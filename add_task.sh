
read -p "Task: " task
read -p "Task Description: " taskdesc
read -p "Business function: " bfunc
read -p "Hours: " hours

echo "$(date +'%d/%m/%Y'),$task,$bfunc,$taskdesc,$hours" >> ./task.csv