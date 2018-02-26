import pandas as pd


class Player:
    ''' Player as defined by squash information

    '''
    numPlayers = 0

    def __init__(self, name, PF=0, PA=0):
        self.name = name
        Player.numPlayers += 1
        self.id = Player.numPlayers
        self.PF = PF
        self.PA = PA

    def get_name(self):
        return self.name

    def get_id(self):
        return self.id

    def add_PF(self, points):
        self.PF += points
        return self.PF

    def get_PF(self):
        return self.PF

    def add_PA(self, points):
        self.PA += points
        return self.PA

    def get_PA(self):
        return self.PA


class Game:
    ''' Game had between two players

    '''
    numGames = 0

    def __init__(self, date, session_num, player1, player2, player1score, player2score):
        self.date = date
        self.session_num = session_num
        self.player1 = player1
        self.player2 = player2
        self.player1score = player1score
        self.player2score = player2score
        Game.numGames += 1
        self.id = Game.numGames

    def get_player1_name(self):
        return self.player1.getname()

    def get_id(self):
        return self.id


class Session:
    ''' Session had consisting of multiple games

    '''

    def __init__(self, date, session_num):
        self.date = date
        self.session_num = session_num
        self.games = {}
        self.players = {}

    def add_game(self, game):
        self.games[game.get_id()] = game
        return self.games[game.get_id()]

    def add_player(self, player):
        self.players[player.get_name()] = player
        return self.players[player.get_name()]

    def get_num_games(self):
        return len(self.games)

    def get_num_players(self):
        return len(self.players)
    
    def get_session_info(self):
        return 'Session {0} was played by {1} players ' \
               'who collectively played {2} games.'.format(self.session_num,
                                                           self.get_num_players(),
                                                           self.get_num_games())


if __name__ == "__main__":
    session_headers = [
        'Player1',
        'Score1',
        'Score2',
        'Player2',
        'Date'
    ]

    session_data = [
        ['Brian', 2, 11, 'Steve', 'll/12/2017'],
        ['James', 2, 11, 'Steve', 'll/12/2017'],
        ['Brian', 12, 10, 'James', 'll/12/2017']
    ]

    session_df = pd.DataFrame(data=session_data, columns=session_headers)

    print(session_df.head())

    # Obtain unique list of players from input table
    players_list = list(set(session_df[['Player1', 'Player2']].stack().values))
    players = {}

    for pl in players_list:
        players[pl] = Player(pl)

    # Define Session object and cycle through games adding game information
    date_of_session = '11/12/2017'
    num_of_session = 3

    session_obj = Session(date_of_session, num_of_session)

    for _, game in session_df.iterrows():
        p1 = players[game['Player1']]
        p2 = players[game['Player2']]

        p1_score = game['Score1']
        p2_score = game['Score2']

        p1.add_PF(p1_score)
        p2.add_PF(p2_score)

        p1.add_PA(p2_score)
        p2.add_PA(p1_score)

        game_obj = Game(
            date_of_session,
            num_of_session,
            p1, p2,
            p1_score, p2_score
        )

        session_obj.add_game(game_obj)
        session_obj.add_player(p1)
        session_obj.add_player(p2)

    print(session_obj.get_session_info())

    for _, playa in players.items():
        print(playa.get_PA())
        print(playa.get_PF())
