import simpy


def car(env):
    print("汽车出发，仿真时间 =", env.now)

    yield env.timeout(5)

    print("汽车到达，仿真时间 =", env.now)


env = simpy.Environment()

env.process(car(env))

env.run()