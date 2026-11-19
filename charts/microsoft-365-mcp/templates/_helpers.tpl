{{- define "m365.name" -}}
{{- .Chart.Name | trunc 63 | trimSuffix "-" -}}
{{- end -}}

{{- define "m365.fullname" -}}
{{- if contains .Chart.Name .Release.Name -}}
{{- .Release.Name | trunc 63 | trimSuffix "-" -}}
{{- else -}}
{{- printf "%s-%s" .Release.Name .Chart.Name | trunc 63 | trimSuffix "-" -}}
{{- end -}}
{{- end -}}

{{- define "m365.selectorLabels" -}}
app.kubernetes.io/name: {{ include "m365.name" . }}
app.kubernetes.io/instance: {{ .Release.Name }}
{{- end -}}

{{- define "m365.labels" -}}
{{ include "m365.selectorLabels" . }}
app.kubernetes.io/version: {{ .Chart.AppVersion | quote }}
app.kubernetes.io/managed-by: {{ .Release.Service }}
helm.sh/chart: {{ printf "%s-%s" .Chart.Name .Chart.Version | replace "+" "_" }}
{{- end -}}

{{- define "m365.image" -}}
{{- if .Values.image.digest -}}
{{ .Values.image.repository }}@{{ .Values.image.digest }}
{{- else -}}
{{ .Values.image.repository }}:{{ .Values.image.tag | default .Chart.AppVersion }}
{{- end -}}
{{- end -}}

{{/* Whether a shared state store is configured. */}}
{{- define "m365.sharedState" -}}
{{- if or .Values.state.url .Values.state.sentinel.hosts -}}true{{- end -}}
{{- end -}}

{{/* Whether an area that moves bytes through links is on (the defaults include several). */}}
{{- define "m365.movesFiles" -}}
{{- $bytes := list "mail" "files" "recordings" "notes" "teams" "excel" "contacts" -}}
{{- if not .Values.areas -}}true{{- else -}}
{{- range .Values.areas -}}{{- if has . $bytes -}}true{{- end -}}{{- end -}}
{{- end -}}
{{- end -}}

{{- define "m365.validate" -}}
{{- $v := .Values -}}
{{- if not $v.key.existingSecret -}}
{{- fail "key.existingSecret is required: the Secret holding the sealing key file (scripts/new-sealing-key.sh writes one)" -}}
{{- end -}}
{{- if include "m365.movesFiles" . -}}
{{- $u := $v.linksBaseUrl -}}
{{- if not (or (hasPrefix "https://" $u) (regexMatch "^http://[a-z0-9.-]+\\.svc(\\.cluster\\.local)?(:[0-9]+)?/?$" $u)) -}}
{{- fail "linksBaseUrl is required while an area that moves files is on: https://…, or http://<service>.<namespace>.svc[.cluster.local][:port]" -}}
{{- end -}}
{{- end -}}
{{- $many := or (gt (int $v.replicaCount) 1) $v.autoscaling.enabled -}}
{{- if and $many (not (include "m365.sharedState" .)) -}}
{{- fail "more than one replica (replicaCount > 1 or autoscaling.enabled) needs state.url or state.sentinel.hosts: without the shared store each pod would grant its own send allowance and know only the file links it issued. For one pod set replicaCount=1" -}}
{{- end -}}
{{- if and $v.state.url $v.state.sentinel.hosts -}}{{- fail "state.url and state.sentinel.hosts are two different stores: set one" -}}{{- end -}}
{{- if and $v.state.url (not (or (hasPrefix "redis://" $v.state.url) (hasPrefix "rediss://" $v.state.url))) -}}{{- fail "state.url is redis://host:port/db or rediss://host:port/db" -}}{{- end -}}
{{- if contains "@" $v.state.url -}}{{- fail "state.url carries credentials: the password goes in state.existingSecret, the user in state.username" -}}{{- end -}}
{{- if and $v.state.url $v.state.tls (hasPrefix "redis://" $v.state.url) -}}{{- fail "state.tls with state.url: use a rediss:// address" -}}{{- end -}}
{{- if and $v.state.existingSecret (not (include "m365.sharedState" .)) -}}{{- fail "state.existingSecret is set but no store is: set state.url or state.sentinel.hosts" -}}{{- end -}}
{{- if and $v.autoscaling.enabled (gt (int $v.autoscaling.minReplicas) (int $v.autoscaling.maxReplicas)) -}}{{- fail "autoscaling.minReplicas is above autoscaling.maxReplicas" -}}{{- end -}}
{{- if has (int $v.metrics.port) (list 8000 8001) -}}{{- fail "metrics.port must differ from 8000 and 8001: metrics are never served on the MCP or the links port" -}}{{- end -}}
{{- if ne $v.service.type "ClusterIP" -}}{{- fail "service.type is ClusterIP: the Service carries the MCP and metrics ports too; publish the links port with ingress.enabled" -}}{{- end -}}
{{- if and $v.ingress.enabled (not (regexMatch "^https://[^/?#]+/?$" $v.linksBaseUrl)) -}}{{- fail "ingress.enabled: linksBaseUrl must be https://<host> with no path, the Ingress serves /d/ and /u/ at its root" -}}{{- end -}}
{{- if not (has $v.logLevel (list "DEBUG" "INFO" "WARNING" "ERROR" "CRITICAL")) -}}{{- fail "logLevel is DEBUG, INFO, WARNING, ERROR or CRITICAL" -}}{{- end -}}
{{- if dig "limits" "cpu" "" $v.resources -}}{{- fail "resources.limits.cpu: no CPU limit; throttling a server that streams files stalls transfers" -}}{{- end -}}
{{- range $k, $_ := $v.podLabels -}}
{{- if has $k (list "app.kubernetes.io/name" "app.kubernetes.io/instance") -}}{{- fail (printf "podLabels.%s is set by the chart: the Deployment selects its pods by it" $k) -}}{{- end -}}
{{- end -}}
{{- if hasKey $v.podAnnotations "microsoft-365-mcp/key-revision" -}}{{- fail "podAnnotations.microsoft-365-mcp/key-revision is set by the chart from key.revision" -}}{{- end -}}
{{- end -}}
