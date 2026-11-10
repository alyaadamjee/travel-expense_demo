#!/bin/zsh
# Rebuilds docs/data.js (the static dashboard's data) from <schema>.spend_unified, load_checks and agency_adjustments.
# usage: DATABRICKS_WAREHOUSE_ID=<id> scripts/export_html_data.sh <catalog.schema>
set -e
S=$1; wh=${DATABRICKS_WAREHOUSE_ID:?set DATABRICKS_WAREHOUSE_ID}; here=${0:A:h}; tmp=$(mktemp -d)
q() {  # run one statement, write tab-separated rows to $2
  perl -e 'my $q=$ARGV[0]; $q=~s/\\/\\\\/g; $q=~s/"/\\"/g; print qq({"warehouse_id":"$ARGV[1]","wait_timeout":"50s","format":"JSON_ARRAY","disposition":"INLINE","statement":"$q"})' "$1" "$wh" > $tmp/q.json
  databricks api post /api/2.0/sql/statements --json @$tmp/q.json | perl -0ne 'die "query failed\n" unless /"state":\s*"SUCCEEDED"/; if(/"data_array":\s*\[(.*)\]\s*\}/s){my $d=$1; while($d=~/\[\s*((?:"(?:[^"\\]|\\.)*"|null|\s|,)*)\s*\]/g){my $r=$1; my @c; while($r=~/("(?:[^"\\]|\\.)*"|null)/g){my $x=$1; $x=~s/^"|"$//g; push @c,$x} print join("\t",@c),"\n"}}' > $2
}
P="CASE WHEN source <> 'Travel agency' THEN NULL WHEN purpose_of_trip IS NULL THEN 'Not given' WHEN purpose_of_trip IN ('Client Meeting','Client Work','Site Visit') THEN 'Client' WHEN purpose_of_trip = 'Internal Meeting' THEN 'Internal' WHEN purpose_of_trip = 'Conference' THEN 'Conferences & events' WHEN purpose_of_trip = 'New Business' THEN 'New business' WHEN purpose_of_trip = 'Training' THEN 'Training' ELSE 'Other' END"
q "SELECT source, category_group, category, employee_name, country, date_format(month,'yyyy-MM'), CAST(amount_usd AS DOUBLE), cc_group_2, cc_group_3, team, $P, line_basis FROM $S.spend_unified" $tmp/rows.tsv
q "SELECT check, CAST(value AS DOUBLE) FROM $S.load_checks" $tmp/checks.tsv
q "SELECT category, action, traveler, CAST(invoice_date AS STRING), detail, country, CAST(original_amount AS DOUBLE), CAST(corrected_amount AS DOUBLE), CAST(amount_kept_usd AS DOUBLE) FROM $S.agency_adjustments ORDER BY original_amount DESC" $tmp/adj.tsv
perl -MJSON::PP -e '
my ($rows,$checks,$adj)=@ARGV; my @cols=qw(source category_group category employee country month amount cc_group_2 department team purpose line_basis);
my (%dict,%D,@out);
open R,"<:utf8",$rows; while(<R>){chomp; my @f=split /\t/,$_,-1;
  push @out,[ map { my $v=$f[$_]; $v=undef if $v eq "null"; $_==6 ? 0+$v : !defined $v ? undef : ($D{$cols[$_]}{$v} //= do { push @{$dict{$cols[$_]}}, $v; $#{$dict{$cols[$_]}} }) } 0..$#cols ] }
open C,"<:utf8",$checks; my @C=map{chomp; my @f=split /\t/; [$f[0],0+$f[1]]}<C>;
open A,"<:utf8",$adj; my @A=map{chomp; my @f=split /\t/,$_,-1; [@f[0..5],0+$f[6],0+$f[7],0+$f[8]]}<A>;
binmode STDOUT,":utf8";
print "// Generated from the demo table spend_unified (fake data). Columns are dictionary-encoded: rows[i][c] indexes dict[col].\n";
print "window.TE_DATA = ", JSON::PP->new->canonical->encode({columns=>\@cols, dict=>\%dict, rows=>\@out, checks=>\@C, adjustments=>\@A}), ";\n";
' $tmp/rows.tsv $tmp/checks.tsv $tmp/adj.tsv > "$here/../docs/data.js"
echo "wrote docs/data.js ($(wc -l < $tmp/rows.tsv | tr -d ' ') rows)"
