create table if not exists projects (
    id bigserial primary key,
    name text not null,
    project_number text default '',
    task_code text default '',
    client text default '',
    site_location text default '',
    borehole_start_date text default '',
    borehole_end_date text default '',
    borehole_include_saturday integer default 0,
    cptu_start_date text default '',
    cptu_end_date text default '',
    cptu_include_saturday integer default 0,
    test_pit_start_date text default '',
    test_pit_end_date text default '',
    test_pit_include_saturday integer default 0,
    geophysics_start_date text default '',
    geophysics_end_date text default '',
    geophysics_include_saturday integer default 0,
    custom_methods_json text default '[]',
    use_borehole integer default 1,
    use_cptu integer default 1,
    use_test_pit integer default 1,
    use_geophysics integer default 1,
    borehole_budget_meters double precision default 0,
    cptu_budget_meters double precision default 0,
    geophysics_budget_meters double precision default 0
);

create table if not exists map_items (
    id bigserial primary key,
    project_id bigint not null references projects(id) on delete cascade,
    item_type text not null,
    item_id text not null,
    geometry_type text not null,
    coords_json text not null,
    location_plan text default '',
    planned_amount double precision default 0,
    status text default 'Planned',
    work_start_date text default '',
    work_end_date text default '',
    notes text default '',
    depth_m double precision default 0,
    exclude_from_history integer default 0
);

create table if not exists import_backups (
    id bigserial primary key,
    project_id bigint not null references projects(id) on delete cascade,
    created_at text not null,
    item_count integer not null,
    backup_json text not null
);

create table if not exists historical_rates (
    id bigserial primary key,
    source_item_id bigint unique,
    project_id bigint not null references projects(id) on delete cascade,
    project_name text default '',
    item_type text not null,
    item_id text not null,
    work_start_date text default '',
    work_end_date text default '',
    completion_month text default '',
    work_days integer default 0,
    depth_m double precision default 0,
    items_per_day double precision default 0,
    meters_per_day double precision,
    recorded_at text not null
);

create index if not exists idx_map_items_project_id on map_items(project_id);
create index if not exists idx_map_items_status on map_items(status);
create index if not exists idx_import_backups_project_id on import_backups(project_id);
create index if not exists idx_historical_rates_project_type_month on historical_rates(project_id, item_type, completion_month);

alter table projects enable row level security;
alter table map_items enable row level security;
alter table import_backups enable row level security;
alter table historical_rates enable row level security;

create policy "authenticated users can read projects" on projects
    for select to authenticated using (true);
create policy "authenticated users can insert projects" on projects
    for insert to authenticated with check (true);
create policy "authenticated users can update projects" on projects
    for update to authenticated using (true) with check (true);
create policy "authenticated users can delete projects" on projects
    for delete to authenticated using (true);

create policy "authenticated users can read map items" on map_items
    for select to authenticated using (true);
create policy "authenticated users can insert map items" on map_items
    for insert to authenticated with check (true);
create policy "authenticated users can update map items" on map_items
    for update to authenticated using (true) with check (true);
create policy "authenticated users can delete map items" on map_items
    for delete to authenticated using (true);

create policy "authenticated users can read import backups" on import_backups
    for select to authenticated using (true);
create policy "authenticated users can insert import backups" on import_backups
    for insert to authenticated with check (true);
create policy "authenticated users can update import backups" on import_backups
    for update to authenticated using (true) with check (true);
create policy "authenticated users can delete import backups" on import_backups
    for delete to authenticated using (true);

create policy "authenticated users can read historical rates" on historical_rates
    for select to authenticated using (true);
create policy "authenticated users can insert historical rates" on historical_rates
    for insert to authenticated with check (true);
create policy "authenticated users can update historical rates" on historical_rates
    for update to authenticated using (true) with check (true);
create policy "authenticated users can delete historical rates" on historical_rates
    for delete to authenticated using (true);
