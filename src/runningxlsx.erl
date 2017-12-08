-module(runningxlsx).

-behaviour(application).

%% Application callbacks
-export([start/2, stop/1]).
-export([tables/0,lookup/2]).
%% ===================================================================
%% Application callbacks
%% ===================================================================

start(_StartType, _StartArgs) ->
    runningxlsx_sup:start_link().

stop(_State) ->
    ok.

lookup(Table,MatchFun)when is_function(MatchFun)->
	case xlsx_holder:current_root() of
		undefined->[];
		RootId->
			case ets:lookup(RootId,Table) of
				[]->[];
				[{_,TabId,_Fields}]-> ets:match(TabId,MatchFun)
		    end
	end;
lookup(Table,Key)->
	case xlsx_holder:current_root() of
		undefined->[];
		RootId->
			case ets:lookup(RootId,Table) of
				[]->[];
				[{_,TabId,_Fields}]-> ets:lookup(TabId,Key)
			end
	end.


tables()->
	case xlsx_holder:current_root() of
		undefined->[];
		RootId-> List = ets:tab2list(RootId),
				lists:map(fun({Table,_,_})-> Table end,List)
	end.