%%%-------------------------------------------------------------------
%%% @author Adrianx Lau <adrianx.lau@gmail.com>
%%% @copyright (C) 2017, 
%%% @doc
%%%
%%% @end
%%% Created : 08. 三月 2017 15:53
%%%-------------------------------------------------------------------
-module(runningxlsx_util).
-author("Adrianx Lau <adrianx.lau@gmail.com>").
-include_lib("kernel/include/file.hrl").
%% API
-export([string_to_term/1,check_list/2,error/2,error/1]).
-export([string_to_float/1,string_to_integer/1,string_to_integer_list/1,string_to_float_list/1]).
-export([get_fileunixtime/1]).
string_to_term(String)->
	case erl_scan:string(String ++ ".") of
		{ok, Tokens, _}->
			erl_parse:parse_term(Tokens);
		Error->
			Error
	end.

check_list(List,integer)->
	lists:all(fun(I)-> is_integer(I) end,List);
check_list(List,atom)->
	lists:all(fun(I)-> is_atom(I) end,List);
check_list(List,string)->
	lists:all(fun(I)-> is_list(I) end,List);
check_list(List,float)->
	lists:all(fun(I)-> is_integer(I) or is_float(I) end,List);
check_list(List,tuple)->
	lists:all(fun(I)-> is_tuple(I)  end,List);
check_list(List,binary)->
	lists:all(fun(I)-> is_binary(I)  end,List);
check_list(List,any)->
	is_list(List).

error(String)->
	erlang:error(String).

error(Format,Args)->
	String = lists:flatten(io_lib:format(Format, Args)),
	erlang:error(String).

string_to_float(String)->
	{ok,Flt} = string_to_term(String),
	if is_integer(Flt) -> {ok,Flt};
		is_float(Flt)-> {ok,Flt};
		true-> {error,Flt}
	end.

string_to_integer(String)->
	{ok,Integer} = string_to_term(String),
	if is_integer(Integer) -> {ok,Integer};
		is_float(Integer)-> {ok,trunc(Integer+0.5)};
		true-> {error,Integer}
	end.
string_to_integer_list(String)->
	{ok,IntList} = string_to_term(String),
	Ints= lists:map(
		fun(Int) when is_integer(Int)-> Int;
		   (Int) when is_float(Int)-> trunc(Int+0.5);
		   (_Int) -> runningxlsx_util:error("string to integer list:~s",[String])
		end,IntList),
	{ok,Ints}.
string_to_float_list(String)->
	{ok,FltList} = string_to_term(String),
	Flts = lists:map(
		fun(Flt) when is_integer(Flt)-> Flt;
			(Flt) when is_float(Flt)-> Flt;
			(_Flt) -> runningxlsx_util:error("string to float list:~s",[String])
		end,FltList),
	{ok,Flts}.

get_fileunixtime(File)->
	case file:read_file_info(File,[{time, universal}]) of
		{ok, FileInfo}->
			LastModify = FileInfo#file_info.mtime,
			calendar:datetime_to_gregorian_seconds(LastModify);
		{error,_Reason}-> {error,0}
	end.