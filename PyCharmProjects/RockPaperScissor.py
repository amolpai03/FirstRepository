import random

def RPS():
    pl_choice = input("Tell me your choice: ")
    options = ("Rock", "Paper", "Scissor")          #Created as a Tuple instead of List as we don't need to change this later
    comp_choice = random.choice(options)

    print("Player chose " + pl_choice + " Computer chose " + comp_choice)
    pl_choice = pl_choice.lower()
    comp_choice = comp_choice.lower()

    #Method 1: Using if else statements----------------------------------------------------------------
    if pl_choice == comp_choice:
        print("Draw")
    elif pl_choice == "rock":
        if comp_choice == "paper":
            print("Computer wins!")
        else:
            print("Player wins!")
    elif pl_choice == "paper":
        if comp_choice == "scissor":
            print("Computer wins!")
        else:
            print("Player wins!")
    elif pl_choice == "scissor":
        if comp_choice == "rock":
            print("Computer wins!")
        else:
            print("Player wins!")
    else:
        print("Wrong input by the Player")
    #------------------------------------------------------------------------------------------
    #Method 2: Using Dictionary----------------------------------------------------------------
    wins = {
            "rock": "scissor",
            "paper": "rock",
            "scissor": "paper"
    }
    if pl_choice in wins:
        if wins[pl_choice] == comp_choice:
            print("Player wins!")
        elif pl_choice == comp_choice:
            print("Draw")
        else:
            print("Computer wins!")
    else:
        print("Wrong input by the Player")

RPS()