-module(runningxlsx).

-behaviour(application).

%% Application callbacks
-export([start/2, stop/1]).
-export([lookup/2,get/3]).
%% ===================================================================
%% Application callbacks
%% ===================================================================

start(_StartType, _StartArgs) ->
    runningxlsx_sup:start_link().

stop(_State) ->
    ok.


lookup(Table,MatchFun)when is_function(MatchFun)->
	[object];
lookup(Table,Key)->
	[object].


get(Table,MatchFun,RetFields)when is_function(MatchFun)->
	[fields];
get(Table,Key,RetFields)->
	[].
