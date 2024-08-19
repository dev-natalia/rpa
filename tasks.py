from robot_script import Robot
from robocorp.tasks import task

@task
def task():
    robot = Robot()
    robot.start_robot(phrase="são paulo", time_period=0)
    
task()