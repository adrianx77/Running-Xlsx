-module(runningxlsx_sup).

-behaviour(supervisor).

%% API
-export([start_link/0]).

%% Supervisor callbacks
-export([init/1]).

%% Helper macro for declaring children of supervisor
-define(CHILD(I, Type), {I, {I, start_link, []}, permanent, 5000, Type, [I]}).
-define(CHILD_SPEC(Id,Mod,Type,Args) , {Id, {Mod, start_link, Args}, permanent, 5000, Type, [Mod]}).
%% ===================================================================
%% API functions
%% ===================================================================

start_link() ->
    supervisor:start_link({local, ?MODULE}, ?MODULE, []).

%% ===================================================================
%% Supervisor callbacks
%% ===================================================================

init([]) ->

    Child = ?CHILD_SPEC(xlsx_holder,xlsx_holder,worker,[]),
    {ok, { {one_for_one, 5, 10}, [Child]} }.

