


# Review of NIST Process to Apply AI and ML to Industrial Robots


----



# Abstract


Approaches categorized as "artificial intelligence" (AI) are enabling significant advances in robotics. Recently, data-centric machine learning has become a prominent tool in a number of disciplines relevant to robotics. AI applied to robotics "can create smarter, faster, cheaper, and more environmentally-friendly production processes that can increase worker productivity, improve product quality, lower costs, and improve worker health and safety. Machine learning algorithms can improve the scheduling of manufacturing processes and reduce inventory requirements." AI's rapid rate of adoption has led to many successes, as well as the need for a measurement science infrastructure to help generate data and qualify it.


For industry to use AI, they must trust what comes out of the AI system. The key idea is to develop data sets and trained AI system, validated through performance evaluation techniques, to allow them to be applied to manufacturing robotic systems. This will allow manufacturers to gain more value from their robots by allowing the robot to "learn" new tasks, and how to better perform existing tasks, without the need for human intervention. NIST is uniquely qualified to address this because of our experience in robot performance characterization, information modeling standards, and robot programming.


**Key words:** robotics, machine learning, artificial intelligence, neural networks, manufacturing


## Reinforcement Learning


Reinforcement Learning (RL) is a machine learning framework for optimizing the behaviour of an agent interacting with an unknown environment [@sutton1998introduction]. Reinforcement Learning enables a robot to autonomously discover an optimal behavior through trial-and-error interactions with its environment [@kormushev2013reinforcement]. Instead of explicitly detailing the solution to a problem, in reinforcement learning the designer of a control task provides feedback in terms of an objective function that measures the one-step performance of the robot.


Reinforcement learning enables a robot to autonomously discover an optimal behavior through trial-and-error interactions with its environment . Instead of explicitly detailing the solution to a problem, in reinforcement learning the designer of a control task provides feedback in terms of an objective function that measures the one-step performance of the robot.


Larouche and Féraud  reinforcement defines reinforcement Learning (RL) to be considered learning through trial and error to control an agent behavior in a stochastic environment: at each time step \[[1](#Reference_1)\]\[[1](#Reference_1)\]\[[2](#Reference_2)\]\[[3](#Reference_3)\], the agent performs an action  <img src="https://latex.codecogs.com/svg.latex?a(t)&space;\in&space;A" title="a(t)&space;\in&space;A" /> , and then perceives from its environment a signal      <img src="https://latex.codecogs.com/svg.latex?o(t)&space;\in&space;\omega" title="o(t)&space;\in&space;\omega" />      called observation, and receives a reward   <img src="https://latex.codecogs.com/svg.latex?t(t)&space;\in&space;R" title="t(t)&space;\in&space;R" />, bounded between   <img src="https://latex.codecogs.com/svg.latex?R_{min}" title="R_{min}" />  and   <img src="https://latex.codecogs.com/svg.latex?R_{max}" title="R_{max}" />. Laroche and Féraud then proposes to share their trajectories expressed in a universal format. A high level definition of the RL algorithms allows to share trajectories between algorithms: a trajectory as a sequence of observations, actions, and rewards can be interpreted by any algorithm in its own decision process and state representation.


Kober  uses training a robot to play table tennis to explain RL concepts. Robot observations of ball position and velocity as well as the internal joint dynamics constitute the state s of the system. The actions a available to the robot could be torque motor commands. A function    <img src="https://latex.codecogs.com/svg.latex?\pi" title="\pi" />   generates the actions based on the state and would be called a policy. This leads to the definition of a reinforcement problem is to find a policy that optimizes the long-term sum of reward    <img src="https://latex.codecogs.com/svg.latex?R(s,a)" title="R(s,a)" />  .








# References


\[1\] <a name="Reference_1"></a>P. Kormushev, S. Calinon,  and D. Caldwell. Reinforcement learning in robotics: Applications and real-world challenges. Robotics, 2(3): pp. 122-148, 2013. 


\[2\] <a name="Reference_2"></a>R. Laroche,  and R. Féraud. Reinforcement Learning Algorithm Selection. arXiv preprint arXiv:1701.08810, 2017. http://citeseerx.ist.psu.edu/viewdoc/download?doi=10.1.1.472.7494&rep=rep1&type=pdf. 


\[3\] <a name="Reference_3"></a>J. Kober, J. A. Bagnell,  and J. Peters. Reinforcement learning in robotics: A survey. The International Journal of Robotics Research, 32(11): pp. 1238-1274, 2013. http://citeseerx.ist.psu.edu/viewdoc/download?doi=10.1.1.910.7004&rep=rep1&type=pdf. 


