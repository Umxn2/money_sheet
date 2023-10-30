

if [ "$(date -v+1d +%d)"   == "1" ]; then
   
    echo "ball" >> /Users/umang/umang/3-1/random/ball.txt
    python /Users/umang/umang/3-1/random/dataset_new.py

fi
