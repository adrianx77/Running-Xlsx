%%%-------------------------------------------------------------------
%%% @author Adrianx Lau <adrianx.lau@gmail.com>
%%% @copyright (C) 2017, 
%%% @doc
%%%
%%% @end
%%% Created : 28. 二月 2017 下午9:04
%%%-------------------------------------------------------------------
-module(xlsx_holder).
-author("Adrianx Lau <adrianx.lau@gmail.com>").
-behaviour(gen_server).

%% API
-export([start_link/1,current_root/0]).

%% gen_server callbacks
-export([init/1,
	handle_call/3,
	handle_cast/2,
	handle_info/2,
	terminate/2,
	code_change/3]).

-define(SERVER, ?MODULE).
-define(CHECK_CONFIG_DATA,{check_config}).
-define(CHECK_INTERVAL,10000).
-define(ROOT_TABLE,'running_root_table').
-define(RUNNING_TABLES_TABLE,'running_tables_ets').
-record(xlsx_field,	{column,name,type,descript}).
-record(xlsx_header,{table,fields,tabid}).
-record(state, {dir,file_list_table}).

%%%===================================================================
%%% API
%%%===================================================================
current_root()->
	case ets:lookup(?ROOT_TABLE,root) of
		[]-> undefined;
		[{root,RootId}]-> RootId
	end.
%%--------------------------------------------------------------------
%% @doc
%% Starts the server
%%
%% @end
%%--------------------------------------------------------------------
-spec(start_link(RXOpt :: term()) ->
	{ok, Pid :: pid()} | ignore | {error, Reason :: term()}).
start_link(Dir) ->
	gen_server:start_link({local, ?SERVER}, ?MODULE, Dir, []).

%%%===================================================================
%%% gen_server callbacks
%%%===================================================================

%%--------------------------------------------------------------------
%% @private
%% @doc
%% Initializes the server
%%
%% @spec init(Args) -> {ok, State} |
%%                     {ok, State, Timeout} |
%%                     ignore |
%%                     {stop, Reason}
%% @end
%%--------------------------------------------------------------------
-spec(init(Args :: term()) ->
	{ok, State :: #state{}} | {ok, State :: #state{}, timeout() | hibernate} |
	{stop, Reason :: term()} | ignore).
init(Dir) ->
	ets:new(?ROOT_TABLE,[set,named_table,protected,{read_concurrency,true}]),
	FilesTime = flush_dir(Dir,[]),
	timer:send_after(?CHECK_INTERVAL,?CHECK_CONFIG_DATA),
	{ok, #state{file_list_table = FilesTime,dir = Dir}}.

%%--------------------------------------------------------------------
%% @private
%% @doc
%% Handling call messages
%%
%% @end
%%--------------------------------------------------------------------
-spec(handle_call(Request :: term(), From :: {pid(), Tag :: term()},
	State :: #state{}) ->
	{reply, Reply :: term(), NewState :: #state{}} |
	{reply, Reply :: term(), NewState :: #state{}, timeout() | hibernate} |
	{noreply, NewState :: #state{}} |
	{noreply, NewState :: #state{}, timeout() | hibernate} |
	{stop, Reason :: term(), Reply :: term(), NewState :: #state{}} |
	{stop, Reason :: term(), NewState :: #state{}}).
handle_call(_Request, _From, State) ->
	{reply, ok, State}.

%%--------------------------------------------------------------------
%% @private
%% @doc
%% Handling cast messages
%%
%% @end
%%--------------------------------------------------------------------
-spec(handle_cast(Request :: term(), State :: #state{}) ->
	{noreply, NewState :: #state{}} |
	{noreply, NewState :: #state{}, timeout() | hibernate} |
	{stop, Reason :: term(), NewState :: #state{}}).
handle_cast(_Request, State) ->
	{noreply, State}.

%%--------------------------------------------------------------------
%% @private
%% @doc
%% Handling all non call/cast messages
%%
%% @spec handle_info(Info, State) -> {noreply, State} |
%%                                   {noreply, State, Timeout} |
%%                                   {stop, Reason, State}
%% @end
%%--------------------------------------------------------------------
-spec(handle_info(Info :: timeout() | term(), State :: #state{}) ->
	{noreply, NewState :: #state{}} |
	{noreply, NewState :: #state{}, timeout() | hibernate} |
	{stop, Reason :: term(), NewState :: #state{}}).
handle_info(?CHECK_CONFIG_DATA,State=#state{dir=Dir,file_list_table = FilesTime})->
	NewFilesTime =flush_dir(Dir,FilesTime),
	timer:send_after(?CHECK_INTERVAL,?CHECK_CONFIG_DATA),
	{noreply,State#state{file_list_table = NewFilesTime}};
handle_info(_Info, State) ->
	{noreply, State}.

%%--------------------------------------------------------------------
%% @private
%% @doc
%% This function is called by a gen_server when it is about to
%% terminate. It should be the opposite of Module:init/1 and do any
%% necessary cleaning up. When it returns, the gen_server terminates
%% with Reason. The return value is ignored.
%%
%% @spec terminate(Reason, State) -> void()
%% @end
%%--------------------------------------------------------------------
-spec(terminate(Reason :: (normal | shutdown | {shutdown, term()} | term()),
	State :: #state{}) -> term()).
terminate(_Reason, _State) ->
	ok.

%%--------------------------------------------------------------------
%% @private
%% @doc
%% Convert process state when code is changed
%%
%% @spec code_change(OldVsn, State, Extra) -> {ok, NewState}
%% @end
%%--------------------------------------------------------------------
-spec(code_change(OldVsn :: term() | {down, term()}, State :: #state{},
	Extra :: term()) ->
	{ok, NewState :: #state{}} | {error, Reason :: term()}).
code_change(_OldVsn, State, _Extra) ->
	{ok, State}.

%%%===================================================================
%%% Internal functions
%%%===================================================================
flush_dir(Dir,OldFilesTime)->
	case check_timeout(Dir,OldFilesTime) of
		false-> OldFilesTime;
		true->
			RootId = ets:new(?RUNNING_TABLES_TABLE, [set, protected,{read_concurrency,true}]),
			Wild = filename:absname_join(Dir,"*.{xlsx,xlsm}"),
			Files = filelib:wildcard(Wild),
			try
				NewFilesTime = lists:foldl(
					fun(File,FilesTime)->
						import_file(File, RootId),
						Time = runningxlsx_util:get_fileunixtime(File),
						[{File,Time}|FilesTime]
					end,[],Files),
				case ets:lookup(?ROOT_TABLE, root) of
					[]->ets:insert(?ROOT_TABLE, {root, RootId});
					[{root, OldRootId}]->ets:insert(?ROOT_TABLE, {root, RootId}), clear_root(OldRootId)
				end,
				NewFilesTime
			catch
			    E:R  ->
				    io:format("~p:~p~n",[E,R]),
				    clear_root(RootId),
				    OldFilesTime
			end
	end.

check_timeout(Dir,OldFilesTime)->
	Wild = filename:absname_join(Dir,"*.{xlsx,xlsm}"),
	Files = filelib:wildcard(Wild),
	NewFileTime = lists:filtermap(
		fun(File)->
			case filename:basename(File) of
				"~$"++_->
					false;
				_Base ->
					FileTime = runningxlsx_util:get_fileunixtime(File),
					{true,{File,FileTime}}
			end
		end,Files),
	Diff = (OldFilesTime--NewFileTime) ++ (NewFileTime -- OldFilesTime),
	if length(Diff) >0 -> true;
		true-> false
	end.

clear_root(RootId)->
	TabList = ets:tab2list(RootId),
	lists:foreach(fun({_TabName,TabId,_})-> ets:delete(TabId) end,TabList),
	ets:delete(RootId).

import_file(File,RootId)->
	LineFun =
		fun(ThisTab,[1|Row],#xlsx_header{table = LastTable,tabid = TabId,fields = FieldList})->
				push_table(RootId,LastTable,TabId,FieldList),
				Fields = process_header([],1,Row),
				TabName = list_to_atom(ThisTab),
				NewTabId = ets:new(TabName,[set,protected,{read_concurrency,true}]),
				NewContext = #xlsx_header{table = TabName,fields = Fields,tabid = NewTabId},
				{next_row,NewContext};
			(ThisTab,[1|Row],_)->
				Fields = process_header([],1,Row),
				TabName = list_to_atom(ThisTab),
				NewTabId = ets:new(TabName,[set,protected,{read_concurrency,true}]),
				NewContext = #xlsx_header{table = TabName,fields = Fields,tabid = NewTabId},
				{next_row,NewContext};
			(_TabName,[2|Row],Context=#xlsx_header{fields = Fields})->
				NewFields = process_header(Fields,2,Row),
				{next_row,Context#xlsx_header{fields = NewFields}};
			(_TabName,[3|Row],Context=#xlsx_header{fields = Fields})->
				NewFields = process_header(Fields,3,Row),
				{next_row,Context#xlsx_header{fields = NewFields}};
			(_TabName,[Line|Row],Context=#xlsx_header{fields = Fields,tabid = TabId})->
				process_data(TabId,Fields,Line,Row),
				{next_row,Context}
		end,
	case xlsx_reader:read(File,undefined,LineFun) of
		#xlsx_header{table = TableName1,tabid = TabId1,fields = FieldList} ->
			push_table(RootId,TableName1,TabId1,FieldList);
		Error-> runningxlsx_util:error("Read:~p~n",[Error])
	end.


push_table(RootId,TableName,TabId,FieldList)->
	ets:insert(RootId,{TableName,TabId,FieldList}).

process_header(_,1,Row)->
	NewRow = lists:zip(lists:seq(1,length(Row)),Row),
	lists:filtermap(
		fun({_,[]})->
			false;
			({Indx,Cell})->
				{true,#xlsx_field{column = Indx,name = Cell}}
		end,NewRow);
process_header(HeaderInfos,2,Row)->
	NewRow = lists:zip(lists:seq(1,length(Row)),Row),
	lists:map(
		fun(#xlsx_field{column = ColIndx}=HeaderInfo)->
			case lists:keyfind(ColIndx,1,NewRow) of
				false->
					HeaderInfo;
				{_,[]}->
					HeaderInfo#xlsx_field{type = string};
				{_,"integer"}->
					HeaderInfo#xlsx_field{type = integer};
				{_,"integer_list"}->
					HeaderInfo#xlsx_field{type = integer_list};
				{_,"float"}->
					HeaderInfo#xlsx_field{type = float};
				{_,"float_list"}->
					HeaderInfo#xlsx_field{type = float_list};
				{_,"tuple"}->
					HeaderInfo#xlsx_field{type = tuple};
				{_,"tuple_list"}->
					HeaderInfo#xlsx_field{type = tuple_list};
				{_,"string"}->
					HeaderInfo#xlsx_field{type = string};
				{_,"string_list"}->
					HeaderInfo#xlsx_field{type = string_list};
				{_,"atom"}->
					HeaderInfo#xlsx_field{type = atom};
				{_,"atom_list"}->
					HeaderInfo#xlsx_field{type = atom_list};
				{_,"binary"}->
					HeaderInfo#xlsx_field{type = binary};
				{_,"binary_list"}->
					HeaderInfo#xlsx_field{type = binary_list};
				{_,"list"}->
					HeaderInfo#xlsx_field{type = list};
				{_,_}->
					HeaderInfo#xlsx_field{type = unknown}
			end
		end,HeaderInfos);
process_header(HeaderInfos,3,Row)->
	NewRow = lists:zip(lists:seq(1,length(Row)),Row),
	lists:map(
		fun(#xlsx_field{column = ColIndx}=HeaderInfo)->
			case lists:keyfind(ColIndx,1,NewRow) of
				false->
					HeaderInfo;
				{_,Cell}->
					HeaderInfo#xlsx_field{descript = Cell}
			end
		end,HeaderInfos).

process_data(_Tab,[],_,_)->
	ok;
process_data(Tab,HeaderInfos,Line,Row)->
	NewRow = lists:zip(lists:seq(1, length(Row)), Row),
	EtsLine = lists:filtermap(
		fun({I, Cell})->
			case lists:keyfind(I, #xlsx_field.column, HeaderInfos) of
				false->false;
				#xlsx_field{type = integer}->
					{true, list_to_integer(Cell)};
				#xlsx_field{type = integer_list}->
					{ok,List} = runningxlsx_util:string_to_integer_list(Cell),
					case runningxlsx_util:check_list(List,integer) of
						true-> {true, List};
						false-> runningxlsx_util:error("invalidate integer list:(~p,~p)",[Line,I])
					end;

				#xlsx_field{type = float}->
					case runningxlsx_util:string_to_float(Cell) of
						{ok,Flt}->{true,Flt};
						{error,_}->runningxlsx_util:error("invalidate float list:(~p,~p)",[Line,I])
					end;
				#xlsx_field{type = float_list}->
					{ok,List} = runningxlsx_util:string_to_float_list(Cell),
					case runningxlsx_util:check_list(List,float) of
						true-> {true, List};
						false-> runningxlsx_util:error("invalidate float list:(~p,~p)",[Line,I])
					end;

				#xlsx_field{type = string}->
					{true, Cell};
				#xlsx_field{type = string_list}->
					{ok,List} = runningxlsx_util:string_to_term(Cell),
					case runningxlsx_util:check_list(List,string) of
						true-> {true, List};
						false-> runningxlsx_util:error("invalidate string list:(~p,~p)",[Line,I])
					end;

				#xlsx_field{type = atom}->
					{true, list_to_atom(Cell)};
				#xlsx_field{type = atom_list}->
					{ok,List} = runningxlsx_util:string_to_term(Cell),
					case runningxlsx_util:check_list(List,atom) of
						true-> {true, List};
						false-> runningxlsx_util:error("invalidate atom list:(~p,~p)",[Line,I])
					end;
				#xlsx_field{type = binary}->
					{true, list_to_binary(Cell)};
				#xlsx_field{type = binary_list}->
					{ok,List} = runningxlsx_util:string_to_term(Cell),
					case runningxlsx_util:check_list(List,binary) of
						true-> {true, List};
						false-> runningxlsx_util:error("invalidate binary list:(~p,~p)",[Line,I])
					end;
				#xlsx_field{type = tuple}->
					{ok,Tuple} = runningxlsx_util:string_to_term(Cell),
					if is_tuple(Tuple)->
						{true, Tuple};
						true->
							runningxlsx_util:error("invalidate tuple :(~p,~p)",[Line,I])
					end;
				#xlsx_field{type = tuple_list}->
					{ok,List} = runningxlsx_util:string_to_term(Cell),
					case runningxlsx_util:check_list(List,tuple) of
						true-> {true, List};
						false-> runningxlsx_util:error("invalidate tuple list:(~p,~p)",[Line,I])
					end;
				#xlsx_field{type = list}->
					{ok,List} = runningxlsx_util:string_to_term(Cell),
					case runningxlsx_util:check_list(List,any) of
						true-> {true, List};
						false-> runningxlsx_util:error("invalidate any list:(~p,~p)",[Line,I])
					end
			end
		end, NewRow),
	ets:insert(Tab, list_to_tuple(EtsLine)).