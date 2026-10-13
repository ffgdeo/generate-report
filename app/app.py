"""Streamlit front end: pick customers / months, generate an editable PPTX via a Databricks job."""
import os
import time
import uuid

import streamlit as st
from databricks.sdk import WorkspaceClient
from databricks.sdk.service.jobs import RunLifeCycleState

CATALOG, SCHEMA = "fd_fevm_classic_catalog", "aibi"
VOLUME_ROOT = os.environ["EXPORT_VOLUME_PATH"]          # /Volumes/<cat>/<schema>/<vol>
JOB_ID = int(os.environ["DECK_JOB_ID"])
WAREHOUSE_ID = os.environ["DATABRICKS_WAREHOUSE_ID"]

st.set_page_config(page_title="Performance Deck Generator", layout="wide")
w = WorkspaceClient()


@st.cache_data(ttl=300)
def query(sql: str):
    r = w.statement_execution.execute_statement(statement=sql, warehouse_id=WAREHOUSE_ID, wait_timeout="30s")
    if r.status.state.value != "SUCCEEDED":
        raise RuntimeError(r.status.error.message if r.status.error else str(r.status.state))
    return r.result.data_array or []


customers = query(f"SELECT customer_id, customer_name, region, segment FROM {CATALOG}.{SCHEMA}.dim_customers ORDER BY region, customer_name")
months = [m[0] for m in query(f"SELECT DISTINCT date_format(order_date, 'yyyy-MM') FROM {CATALOG}.{SCHEMA}.fact_orders ORDER BY 1")]
cust_label = {int(c[0]): f"{c[1]}  ({c[2]} / {c[3]})" for c in customers}

st.title("Performance Deck Generator")
st.caption("Pick a selection, get an editable PowerPoint. Replaces the Power BI button + Power Automate flow.")

left, right = st.columns([2, 1])
with left:
    sel_customers = st.multiselect("Customers (empty = all)", list(cust_label), format_func=cust_label.get)
    sel_months = st.multiselect("Months (empty = all)", months, default=months[-1:])
with right:
    exclude_pending = st.toggle("Exclude pending orders", value=False)
    regions = sorted({c[2] for c in customers if not sel_customers or int(c[0]) in sel_customers})
    st.info(f"Regions in selection: **{', '.join(regions)}**\n\n"
            + ("Slide set: **multi-region**" if len(regions) > 1 else f"Slide set: **{regions[0]}**"))

if st.button("Generate deck", type="primary"):
    request_id = uuid.uuid4().hex[:12]
    out_path = f"{VOLUME_ROOT}/{request_id}/Performance_Deck.pptx"
    user = st.context.headers.get("X-Forwarded-Email", "unknown")
    run = w.jobs.run_now(job_id=JOB_ID, job_parameters={
        "customer_ids": ",".join(map(str, sel_customers)),
        "months": ",".join(sel_months),
        "exclude_pending": str(exclude_pending).lower(),
        "requested_by": user,
        "output_path": out_path,
    })
    with st.status("Building deck on serverless compute...", expanded=True) as status:
        while True:
            r = w.jobs.get_run(run.run_id)
            state = r.state.life_cycle_state
            status.update(label=f"Job state: {state.value}")
            if state in (RunLifeCycleState.TERMINATED, RunLifeCycleState.SKIPPED, RunLifeCycleState.INTERNAL_ERROR):
                break
            time.sleep(4)
        if r.state.result_state and r.state.result_state.value == "SUCCESS":
            status.update(label="Deck ready", state="complete")
            st.session_state["deck"] = w.files.download(out_path).contents.read()
            st.session_state["run_url"] = r.run_page_url
        else:
            status.update(label="Job failed", state="error")
            st.session_state.pop("deck", None)
            st.error(f"{r.state.state_message or 'See run for details'}  \n{r.run_page_url}")

if "deck" in st.session_state:
    st.download_button("Download PPTX", st.session_state["deck"], file_name="Performance_Deck.pptx",
                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    st.caption(f"Job run: {st.session_state['run_url']}")
