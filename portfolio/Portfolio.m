classdef Portfolio < handle
    %   Portfolio optimization class used by portfoliotool
    %   
    %   Copyright 2009-2011 The MathWorks, Inc

    properties (Access = protected)
        % Imported data
        prices         = [];      % Prices series
        benchmark      = [];      % Benchmark series
        dates          = [];      % Dates series
        priceslabels   = [];      % Prices series labels 
        benchmarklabel = [];      % Benchmark label
        % Asset selection
        assetselection = [];      % Asset selection vector (logical)
        % Return series
        decayfactor    = 1;       % Decay factor for weighted returns
        uselogrets     = false;   % True for continuously compounded returns
        % Optimization results
        pf_rsk         = [];      % Portfolio risks
        pf_ret         = [];      % Portfolio returns
        pf_weights     = [];      % Portfolio weights
        % Business days assumtion
        businessdayspermonth = 21;
        businessdaysperyear  = 21*12;  % 252 days
    end
    
   
    methods (Access = public)
        
        function this = Portfolio()
        % Constructor
        end
        
        % Getter methods
        function prices = getPrices(this),                 prices = this.prices(:,this.assetselection);                 end
        function benchmark = getBenchmark(this),           benchmark = this.benchmark;                                  end
        function dates = getDates(this),                   dates = this.dates;                                          end
        function priceslabels = getPricesLabels(this),     priceslabels = this.priceslabels(this.assetselection);       end
        function benchmarklabel = getBenchmarkLabel(this), benchmarklabel = this.benchmarklabel;                        end
        function assetselection = getAssetSelection(this), assetselection = this.assetselection;                        end

        function returns = getReturnSeries(this)
        % Compute return series
        
            % Returns are depending on decay factor and compounding option
            % Call helper function
            returns = this.computeReturnSeries(this.prices(:,this.assetselection));
        end
      
        function [exp_ret,exp_rsk,exp_covariance,annualized_ret,annualized_rsk] = getStatistics(this,all_assets)
        % Compute daily expected risk/return and covariance, and annualized risk/return of selected assets
        % - if input "all_assets" is defined and true, return stats of complete data set

            % compute unweighted return series
            if exist('all_assets','var') && all_assets == true
                returns = this.computeReturnSeries(this.prices,1);
            else
                returns = this.computeReturnSeries(this.prices(:,this.assetselection),1);
            end
            % get stats using ewstats
            [exp_ret,exp_covariance] = ewstats(returns,this.decayfactor);
            exp_rsk = sqrt(diag(exp_covariance));
            exp_rsk = exp_rsk(:)';  % row vector
            % annualized return and volatility
            annualized_rsk = exp_rsk./sqrt(1/this.businessdaysperyear);
            annualized_ret = exp_ret*this.businessdaysperyear;
        end

        function [exp_ret,exp_rsk,annualized_ret,annualized_rsk] = getBenchmarkStatistics(this)
        % Compute daily expected return/risk of benchmark series

            if isempty(this.benchmark)
                exp_ret = [];
                exp_rsk = [];
                annualized_rsk = [];
                annualized_ret = [];
                return
            end
            % compute unweighted benchmark return series
            benchmarkreturns = this.computeReturnSeries(this.benchmark,1);
            % compute stats using ewstats
            [exp_ret,exp_variance] = ewstats(benchmarkreturns,this.decayfactor);
            exp_rsk = sqrt(exp_variance);
            % annualized return and volatility
            annualized_rsk = exp_rsk./sqrt(1/this.businessdaysperyear);
            annualized_ret = exp_ret*this.businessdaysperyear;
        end
        
        function enableAsset(this,assetnumber,state)
        % Enable/disable single asset for optimzation
            if (assetnumber > 0) && (assetnumber <= length(this.assetselection)) && islogical(state)
                this.assetselection(assetnumber) = state;
            end
        end
        
        function useLogReturns(this,val)
        % Select logarithmic/simple return series
           this.uselogrets = val;
        end
        
        function setDecayFactor(this,val)
        % Set new decay factor
            this.decayfactor = val;
        end
        
        function importData(this,prices,benchmark,dates,priceslabels,benchmarklabel)
        % Import new data set

            % reset portfolio
            resetPortfolio(this);
            % input handling
            if nargin < 6
                benchmarklabel = [];
            end
            if nargin < 5
                priceslabels = [];
            end
            if nargin < 4
                dates = [];
            end
            if nargin < 3
                benchmark = [];
            end
            if nargin < 2
                prices = [];
            end
            % assign new data
            this.prices         = prices;
            this.assetselection = true(1,size(prices,2));  % use all assets by default
            this.benchmark      = benchmark;
            this.dates          = dates;
            this.priceslabels   = priceslabels;
            this.benchmarklabel = benchmarklabel;
        end
        
        function error_msg = computeEfficientFrontier(this,constraintSet)
        % Compute efficient frontier
        
            % constraintSet:  optional constraints matrix (e.g. from portcons)
            % error_msg:      contains error message in case of failure, empty otherwise

            % predefine output
            error_msg = [];
            
            % only run if data available
            if isempty(this.prices(:,this.assetselection))
                error_msg = 'No data available';
                return
            end
            
            % get expected returns and covariance matrix
            [eret,~,ecov] = getStatistics(this);
            
            % compute efficient frontier with 15 portfolios
            try
                [rsk,ret,weights] = portopt(eret,ecov,15,[],constraintSet);
            catch ME
                % portopt failed
                error_msg = ME.message;
                rsk = [];
                ret = [];
                weights = [];
            end
            
            % assign results
            this.pf_rsk       = rsk;     % Portfolio risks
            this.pf_ret       = ret;     % Portfolio returns
            this.pf_weights   = weights; % Portfolio weights
            
        end
        
        function [pf_ret, pf_rsk, pf_weights, annualized_ret, annualized_rsk] = getOptimizationResults(this)
        % Get optimization results
        
            pf_ret = this.pf_ret;
            pf_rsk = this.pf_rsk;
            pf_weights = this.pf_weights;
            % annualized return and volatility
            annualized_rsk = pf_rsk./sqrt(1/this.businessdaysperyear);
            annualized_ret = pf_ret*this.businessdaysperyear;
        end
        
        function metrics = getPerformanceMetrics(this,pf_number,riskfreerate)
        % Get performance metrics

            % pf_number:     Portfolio number
            % riskfreerate:  Risk-free rate
            %
            % metrics:       Structure containing following fields:
            %                  annualizedvolatility
            %                  annualizedreturn
            %                  correlation
            %                  sharperatio
            %                  alpha
            %                  riskadjusted_return
            %                  inforatio
            %                  trackingerror
            %                  maxdrawdown

            % check input
            if (pf_number <= 0) || (pf_number > length(this.pf_ret)) || isempty(riskfreerate)
                metrics = [];
                return
            end
            
            % get optimization results
            weights = this.pf_weights(pf_number,:);
            weights = weights(:);   % use column vectors
            pf_prices = this.prices(:,this.assetselection)*weights;  % portfolio prices

            % compute portfolio/index return series (weighted if defined)
            pf_returns = this.computeReturnSeries(pf_prices);            
            b_returns  = this.computeReturnSeries(this.benchmark);
            
            % Annualized return and volatility
            metrics.annualizedvolatility = this.pf_rsk(pf_number)/sqrt(1/this.businessdaysperyear);
            metrics.annualizedreturn = this.pf_ret(pf_number)*this.businessdaysperyear;
            
            % Correlation
            if ~isempty(b_returns)
                metrics.correlation = corrcoef([pf_prices(:),this.benchmark(:)]);
                metrics.correlation = metrics.correlation(1,2);
            else
                metrics.correlation = '-';
            end
            
            % Sharpe ratio
            metrics.sharperatio = sharpe(pf_returns, riskfreerate);

            % Alpha, risk-adjusted return
            if ~isempty(b_returns)
                [alpha, ra_return]  = portalpha(pf_returns, b_returns, riskfreerate,'capm');
            else
                alpha = '-';
                ra_return = '-';
            end
            metrics.alpha = alpha;
            metrics.riskadjusted_return = ra_return;

            % Info ratio, tracking error
            if ~isempty(b_returns)
                [ir,te] = inforatio(pf_returns, b_returns);
            else
                ir = '-';
                te = '-';
            end
            metrics.inforatio = ir;
            metrics.trackingerror = te;

            % Max drawdown
            metrics.maxdrawdown = maxdrawdown(pf_returns,'arithmetic');

        end
        
        function VaR = computeHistoricalVaR(this,pf_number,confidence_level,axes_handle)
        % Compute and visualize historical Value at Risk
            
            % pf_number:            Portfolio number
            % confidence_level      Confidence level (default 0.95)
            % axes_handle           (optional) Visualize result to this graphics handle
            %
            % VaR                   Value at Risk (monthly)

            % handle inputs
            if (pf_number <= 0) || (pf_number > length(this.pf_ret))
                VaR = [];
                return
            end
            if nargin < 3
                confidence_level = 0.95;
            end
            
            % get optimization results and compute per
            weights = this.pf_weights(pf_number,:);
            weights = weights(:);   % use column vectors
            pf_prices = this.prices(:,this.assetselection)*weights;  % portfolio prices
            % compute monthly portfolio returns
            % Note: we don't weight returns for our VaR analysis
            monthly_pf_returns = this.computeMonthlyReturns(pf_prices);
            % do we have enough data points?
            if isempty(monthly_pf_returns)
                VaR = []; 
            else
                % use percentage
                monthly_pf_returns = monthly_pf_returns * 100;
                % Sort returns from smallest to largest
                sorted_returns = sort(monthly_pf_returns);
                % Store the number of returns
                num_returns = numel(monthly_pf_returns);
                % Calculate the index of the sorted return that will be VaR
                VaR_index = ceil((1-confidence_level)*num_returns);
                % Use the index to extract VaR from sorted returns
                VaR = sorted_returns(VaR_index);
            end
            
            % Plot results if requested
            if exist('axes_handle','var') && ishandle(axes_handle)
                % make this axes current
                axes(axes_handle);
                hold('off');
                if isempty(VaR)
                    % Show message
                    cla('reset');
                    axis('off');
                    text(0.2,0.5,{'Too few observations', '  to calculate VaR'});
                else
                    axis('on');
                    % Histogram data
                    [count,bins] = hist(monthly_pf_returns,30);
                    x_min = min(monthly_pf_returns);
                    x_max = max(monthly_pf_returns);
                    x_lim = max(abs([x_min,x_max]));
                    % Create 2nd data set that is zero above Var point
                    count_cutoff = count.*(bins < VaR);
                    % Scale bins
                    scale = (bins(2)-bins(1))*num_returns;
                    % Plot full data set
                    bar(bins,count/scale,'FaceColor',[0.1,0.5,1]);
                    set(axes_handle,'XLim',[-x_lim,x_lim]);
                    hold('on');
                    % Plot cutoff data set
                    bar(bins,count_cutoff/scale,'FaceColor',[1,0.2,0]);
                    grid('on');
                    hold('off');
                    box('off');
                    set(axes_handle,'YTickLabel',[]);
                    title(['Monthly Value at Risk: ',sprintf('%2.2f',VaR),'%'],'FontSize',9);
                end
            end
        end

        function VaR = computeParameticVaR(this,pf_number,confidence_level,axes_handle)
        % Compute and visualize parametric Value at Risk
            
            % pf_number:            Portfolio number
            % confidence_level      Confidence level (default 0.95)
            % axes_handle           (optional) Visualize result to this graphics handle
            %
            % VaR                   Value at Risk (monthly)

            % handle inputs
            if (pf_number <= 0) || (pf_number > length(this.pf_ret))
                VaR = [];
                return
            end
            if nargin < 3
                confidence_level = 0.95;
            end
            
            % get optimization results and compute per
            weights = this.pf_weights(pf_number,:);
            weights = weights(:);   % use column vectors
            pf_prices = this.prices(:,this.assetselection)*weights;  % portfolio prices

            % compute monthly portfolio returns
            % Note: we don't weight returns for our VaR analysis
            monthly_pf_returns = this.computeMonthlyReturns(pf_prices);
            % do we have enough data points?
            if isempty(monthly_pf_returns)
                VaR = []; 
            else
                % use percentage
                monthly_pf_returns = monthly_pf_returns * 100;

                % Calculate mean and standard deviation of returns
                [mu,sigma] = normfit(monthly_pf_returns);
                % Calculate VaR Estimate using Normal Distribution Fit
                VaR = sigma*norminv(1-confidence_level) + mu;
            end
            
            % Plot results if requested
            if exist('axes_handle','var') && ishandle(axes_handle)
                % make this axes current
                axes(axes_handle);
                hold('off');
                if isempty(VaR)
                    % Show message
                    cla('reset');
                    axis('off');
                    text(0.2,0.5,{'Too few observations', '  to calculate VaR'});
                else
                    axis('on');
                    % Construct domain of probabilities to graph distribution.
                    % 100 equally spaced points between min and max return
                    x_lim = max(abs(monthly_pf_returns));
                    x_min = -abs(x_lim);
                    x_max = abs(x_lim);
                    x_full = linspace(x_min,x_max,100);
                    x_partial = x_full(x_full < VaR);
                    y_full = normpdf(x_full,mu,sigma);
                    y_partial = normpdf(x_partial,mu,sigma);
                    area(x_full,y_full,'FaceColor',[0.1,0.5,1]);
                    hold('on');
                    if ~isempty(x_partial)
                        area(x_partial,y_partial,'FaceColor',[1,0.2,0]);    
                    end
                    grid('on');
                    % Histogram data
                    [count,bins] = hist(monthly_pf_returns,30);
                    % Scale bins
                    num_returns = numel(monthly_pf_returns);
                    scale = (bins(2)-bins(1))*num_returns;
                    % Plot full data set
                    a = bar(bins,count/scale,'w');
                    set(axes_handle,'XLim',[-x_lim,x_lim]);
                    set(get(a,'Children'),'FaceAlpha',0.2)
                    box('off');
                    hold('off');
                    set(axes_handle,'YTickLabel',[]);
                    title(['Monthly Value at Risk: ',sprintf('%2.2f',VaR),'%'],'FontSize',9);
                end
            end
        end        
    end
       
    methods (Access = protected)

        function resetPortfolio(this)
        % Reset all data and parameters
            
            % reset all
            this.prices         = [];   % Prices series
            this.benchmark      = [];   % Benchmark series
            this.dates          = [];   % Dates series
            this.priceslabels   = [];   % Prices series labels
            this.benchmarklabel = [];   % Benchmark label        
            this.assetselection = [];   % Asset selection vector (logical)
            this.pf_rsk         = [];   % Portfolio risks
            this.pf_ret         = [];   % Portfolio returns
            this.pf_weights     = [];   % Portfolio weights
            this.decayfactor    = 1;
            this.uselogrets     = false;
        end 

        function rets = computeReturnSeries(this,series,decayfactor)
        % Return exponentially weighted return series
        
            % only run if input not null
            if isempty(series)
                rets = [];
                return
            end
            % if no decay factor specified, use object member
            if ~exist('decayfactor','var')
                decayfactor = this.decayfactor;
            end
            % Define observation dates
            if isempty(this.dates)
                % create default
                dates = (1:size(this.prices,1))';
            else
                dates = this.dates;
            end
            % Compute return series
            if this.uselogrets
                rets = tick2ret(series,dates,'Continuous');
            else
                rets = tick2ret(series,dates,'Simple');
            end
            % Weighted returns
            if decayfactor ~= 1
                retwts =  (decayfactor).^(size(rets,1)-1:-1:0)';
                rets = rets.*repmat(retwts,1,size(rets,2));
            end

        end
        
        
        function monthly_returns = computeMonthlyReturns(this,prices)
        % Compute monthly return series
        
            % we need at least 30 observations
            if length(prices) < 30
                monthly_returns = [];
            else
                if this.uselogrets
                    monthly_returns = log(prices(this.businessdayspermonth+1:end)./prices(1:end-this.businessdayspermonth));
                else
                    monthly_returns = prices(this.businessdayspermonth+1:end)./prices(1:end-this.businessdayspermonth) - 1;
                end
            end
        end
        
    end
    
end

