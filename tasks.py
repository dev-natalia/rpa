from RPA.Robocorp.WorkItems import task
from robot_script import Robot

@task
def task():
    robot = Robot()
    robot.start_robot(phrase="são paulo", time_period=0)
    
