#
# PROMPT
#

PROMPT_BRACKET_BEGIN='%{$fg_bold[reset_color]%}['

# host: not bold, standard terminal color
PROMPT_HOST='%{$reset_color%}%m'

PROMPT_SEPARATOR='%{$reset_color%}:'

# dir: previous_dir/current_dir, bold, standard terminal color
PROMPT_DIR='%{$reset_color%}%B%2~%b'

PROMPT_BRACKET_END='%{$fg_bold[reset_color]%}]'

# user: bold green
PROMPT_USER='%{$fg_bold[green]%}%n%{$reset_color%}'

PROMPT_SIGN='%{$reset_color%}%#'

GIT_PROMPT_INFO='$(git_prompt_info)'

PROMPT="${PROMPT_BRACKET_BEGIN}${PROMPT_HOST}${PROMPT_SEPARATOR}${PROMPT_DIR}${PROMPT_BRACKET_END}${GIT_PROMPT_INFO}
${PROMPT_BRACKET_BEGIN}${PROMPT_USER}${PROMPT_BRACKET_END}${PROMPT_SIGN} "

#
# Git repository
#

# git: bold green
ZSH_THEME_GIT_PROMPT_PREFIX="%{$reset_color%} on %{$fg[green]%}"
ZSH_THEME_GIT_PROMPT_SUFFIX="%{$reset_color%}"
ZSH_THEME_GIT_PROMPT_CLEAN=''