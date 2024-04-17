#!/bin/bash

Xvfb :99 &
XVFB_PID=$!

sleep 1

export DISPLAY=:99

"$@"

kill -9 $XVFB_PID
