#!/bin/zsh
# Creates the Lakeview (AI/BI) dashboard in your workspace from dashboard/travel_expense.lvdash.json.
# usage: DATABRICKS_WAREHOUSE_ID=<id> scripts/deploy_dashboard.sh <catalog.schema> "/Users/<you>/<folder>"
# Needs the databricks CLI (logged in) and perl. Prints the new dashboard id.
set -e
schema=$1; folder=$2; wh=${DATABRICKS_WAREHOUSE_ID:?set DATABRICKS_WAREHOUSE_ID}
here=${0:A:h}
perl -MJSON::PP -e '
  local $/; open my $f, "<", $ARGV[0] or die; my $s = <$f>; $s =~ s/\{\{schema\}\}/$ARGV[1]/g;
  my $d = JSON::PP->new->decode($s);
  print JSON::PP->new->encode({display_name => "Travel & Expense 2026 YTD (demo)", warehouse_id => $ARGV[2],
                               parent_path => $ARGV[3], serialized_dashboard => JSON::PP->new->encode($d)});
' "$here/../dashboard/travel_expense.lvdash.json" "$schema" "$wh" "$folder" > "${TMPDIR:-/tmp}/te_dash.json"
databricks lakeview create --json @"${TMPDIR:-/tmp}/te_dash.json" -o json | perl -ne 'print "dashboard_id: $1\n" if /"dashboard_id":\s*"([^"]+)"/'
